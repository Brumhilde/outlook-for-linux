import { _electron as electron } from 'playwright';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Shared scaffolding for the E2E specs. Keeps the launch/discover/cleanup
// boilerplate in one place so each spec can stay focused on its assertions.

// Hosts the main window can be on before sign-in: Outlook itself, or the
// Microsoft login page it redirects to.
export const MAIN_WINDOW_HOSTNAMES = new Set([
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.cloud.microsoft',
  'outlook.live.com',
  'login.microsoftonline.com',
  'login.live.com',
]);

/**
 * Launch the app with an isolated userData dir and the supplied config.
 *
 * Pass `allowEval: true` only when the test needs `electronApp.evaluate`;
 * it sets `E2E_TESTING=true`, which disables the main window's
 * `globalThis.eval` defense (`browserWindowManager.js`). Default is
 * `false` so a launch from this helper stays environmentally identical to
 * a real user launch unless the test opts in.
 *
 * @param {{ prefix: string, config?: object, allowEval?: boolean }} options
 * @returns {Promise<{ electronApp: import('playwright').ElectronApplication, userDataDir: string }>}
 */
export async function startApp({ prefix, config, allowEval = false }) {
  const userDataDir = mkdtempSync(join(tmpdir(), prefix));
  if (config) {
    writeFileSync(join(userDataDir, 'config.json'), JSON.stringify(config));
  }
  const env = {
    ...process.env,
    E2E_USER_DATA_DIR: userDataDir,
  };
  if (allowEval) {
    env.E2E_TESTING = 'true';
  }
  const electronApp = await electron.launch({
    args: [
      './app/index.js',
      ...(process.env.CI ? ['--no-sandbox'] : []),
    ],
    env,
    timeout: 30000,
  });
  await electronApp.firstWindow({ timeout: 30000 });
  // Give the app a moment to create all windows and start loading.
  await new Promise((resolve) => setTimeout(resolve, 4000));
  return { electronApp, userDataDir };
}

/**
 * Find the main Outlook window in an electronApp by hostname match. Returns
 * undefined if no window has navigated to an Outlook or Microsoft login URL.
 */
export function findMainOutlookWindow(electronApp) {
  return electronApp.windows().find((w) => {
    const url = w.url();
    try {
      return MAIN_WINDOW_HOSTNAMES.has(new URL(url).hostname);
    } catch {
      return false;
    }
  });
}

/**
 * Wait until the main window has redirected to login.microsoftonline.com.
 */
export async function waitForLoginRedirect(mainWindow) {
  await mainWindow.waitForLoadState('load', { timeout: 30000 });
  await mainWindow.waitForURL(
    (url) => {
      try {
        return new URL(url).hostname === 'login.microsoftonline.com';
      } catch {
        return false;
      }
    },
    { timeout: 30000 }
  );
}

/**
 * Best-effort shutdown of the electronApp + temp userData dir. Mirrors
 * the cleanup pattern used by `tests/e2e/smoke.spec.js`: graceful close
 * with a 5-second timeout, then SIGKILL fallback, then rmSync.
 */
export async function closeAndCleanup({ electronApp, userDataDir }) {
  if (electronApp) {
    try {
      const closePromise = electronApp.close();
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Close timeout')), 5000)
      );
      await Promise.race([closePromise, timeoutPromise]);
    } catch (error) {
      console.debug('Process cleanup:', error.message);
      try {
        const pid = electronApp.process()?.pid;
        if (pid) {
          process.kill(pid, 'SIGKILL');
        }
      } catch {
        // Ignore
      }
    }
  }
  if (userDataDir) {
    try {
      rmSync(userDataDir, { recursive: true, force: true });
    } catch (e) {
      console.warn(
        `Could not remove temp directory ${userDataDir}: ${e.message}`
      );
    }
  }
}
