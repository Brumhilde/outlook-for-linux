const { defineConfig } = require('@playwright/test');

// Session directories contain live Microsoft auth tokens stored in plaintext
// (--password-store=basic). Never commit session data or use a path outside
// a gitignored directory (e.g. .auth/, which .gitignore already covers).
if (!process.env.E2E_SESSION_DIR) {
  throw new Error(
    'E2E_SESSION_DIR must be set to a gitignored directory containing a ' +
    'logged-in session. Create one by running the app once with ' +
    'E2E_USER_DATA_DIR=<that directory> and signing in to Outlook.'
  );
}

module.exports = defineConfig({
  testDir: './tests/e2e/authenticated',
  timeout: 90000,
  fullyParallel: false,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  workers: 1,
  reporter: 'list',
  use: {
    trace: 'on-first-retry',
    sessionDir: process.env.E2E_SESSION_DIR,
  },
});
