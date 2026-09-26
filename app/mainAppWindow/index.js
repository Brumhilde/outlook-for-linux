const {
  shell,
  BrowserWindow,
  app,
  dialog,
  nativeImage,
} = require("electron");
const login = require("../login");
const customCSS = require("../customCSS");
const Menus = require("../menus");
const { SpellCheckProvider } = require("../spellCheckProvider");
const { execFile } = require("node:child_process");
const TrayIconChooser = require("../browser/tools/trayIconChooser");
require("../appConfiguration");
const ConnectionManager = require("../connectionManager");
const ssoPasswordPrefill = require("../ssoPasswordPrefill");
const BrowserWindowManager = require("../mainAppWindow/browserWindowManager");
const mailto = require("../mailto");
const { isOutlookHost, isOutlookUrl } = require("../helpers/outlookHosts");
const { POPOUT_WINDOW_ARG } = require("../helpers/popout");
const os = require("node:os");
const path = require("node:path");

let iconChooser;
let intune;
let isControlPressed = false;
// Counter for tracking about:blank navigation attempts to handle authentication flows.
// Microsoft 365 web apps sometimes navigates to about:blank during SSO/auth redirects, and we need to
// intercept these and handle them in a hidden window to complete the auth process.
let aboutBlankRequestCount = 0;
let config;
let window = null;
let appConfig = null;
let connectionManager = null;
let menus = null;

const isMac = os.platform() === "darwin";

// Microsoft Cloud App Security proxy suffix. Tenants that route Outlook
// through Defender for Cloud Apps (MCAS) load and store cookies at
// `*.mcas.ms` rather than the underlying Microsoft domain. Strip the
// suffix before matching against AUTH_DOMAINS / AUTH_LOGIN_DOMAINS so the
// proxied flavour is treated the same as the canonical hostname.
const MCAS_SUFFIX = '.mcas.ms';
function stripMcasSuffix(hostname) {
  return hostname.endsWith(MCAS_SUFFIX)
    ? hostname.slice(0, -MCAS_SUFFIX.length)
    : hostname;
}

// Microsoft auth domains whose cookies should be checked/cleaned
const AUTH_DOMAINS = [
  'login.microsoftonline.com',
  'login.microsoft.com',
  'outlook.cloud.microsoft',
  'microsoft.com',
  'office.com',
  'office365.com',
  'live.com',
  'microsoftonline.com',
];

// Azure AD / MSAL / SharePoint auth cookie names
const AUTH_COOKIE_NAMES = new Set([
  'ESTSAUTH',
  'ESTSAUTHPERSISTENT',
  'ESTSAUTHLIGHT',
  'SignInStateCookie',
  'AADSSO',
  'buid',
  'fpc',
  'x-ms-gateway-slice',
  'stsservicecookie',
  'CCState',
  'FedAuth',
  'rtFa',
  'msal.cache.encryption',
]);

// Auth cookies preserved during force-clean recovery so the Microsoft
// account chooser stays prefilled after session expiry (issue #2364).
const PRESERVE_ON_RECOVERY = new Set(['ESTSAUTHPERSISTENT']);

// localStorage key patterns for MSAL auth tokens
const AUTH_LOCAL_STORAGE_PATTERNS = [
  'refresh_token', 'msal.token', 'msal.',
  'EncryptionKey', 'authSessionId', 'LogoutState',
  'accessToken', 'idtoken', 'Account', 'Authority', 'ClientInfo',
];

/**
 * Checks session cookies for Microsoft auth domains and removes expired
 * or removable auth cookies. When forceCleanAll is true, removes all auth
 * cookies except those in PRESERVE_ON_RECOVERY (e.g. ESTSAUTHPERSISTENT)
 * so the next interactive login can still show the remembered-account banner.
 * @returns {{ cleaned: number, total: number, expired: number }}
 */
async function cleanExpiredAuthCookies(windowSession, forceCleanAll = false) {
  try {
    const allCookies = await windowSession.cookies.get({});
    const nowSeconds = Date.now() / 1000;

    const authCookies = allCookies.filter(cookie => {
      const domain = stripMcasSuffix((cookie.domain || '').replace(/^\./, ''));
      const isAuthDomain = AUTH_DOMAINS.some(d => domain === d || domain.endsWith('.' + d));
      return isAuthDomain && AUTH_COOKIE_NAMES.has(cookie.name);
    });

    const expired = authCookies.filter(c => c.expirationDate && c.expirationDate < nowSeconds);
    const cookiesToRemove = forceCleanAll
      ? authCookies.filter(c => !PRESERVE_ON_RECOVERY.has(c.name))
      : expired;

    if (cookiesToRemove.length === 0) {
      console.debug('[AUTH_RECOVERY] Cookie check:', { total: authCookies.length, expired: expired.length });
      return { cleaned: 0, total: authCookies.length, expired: expired.length };
    }

    console.info('[AUTH_RECOVERY] Cleaning auth cookies:', {
      mode: forceCleanAll ? 'force-all' : 'expired-only',
      removing: cookiesToRemove.length,
      total: authCookies.length,
    });

    const results = await Promise.all(cookiesToRemove.map(async (cookie) => {
      try {
        const protocol = cookie.secure ? 'https' : 'http';
        const domain = cookie.domain.startsWith('.') ? cookie.domain.substring(1) : cookie.domain;
        const url = `${protocol}://${domain}${cookie.path || '/'}`;
        await windowSession.cookies.remove(url, cookie.name);
        return true;
      } catch (err) {
        console.warn('[AUTH_RECOVERY] Failed to remove cookie:', { name: cookie.name, error: err.message });
        return false;
      }
    }));
    const removedCount = results.filter(Boolean).length;

    console.info(`[AUTH_RECOVERY] Cleaned ${removedCount}/${cookiesToRemove.length} auth cookies`);
    return { cleaned: removedCount, total: authCookies.length, expired: expired.length };
  } catch (error) {
    console.error('[AUTH_RECOVERY] Cookie check failed:', error.message);
    return { cleaned: 0, total: 0, expired: 0 };
  }
}

// Some localStorage tokens get encrypted by a Session cookie 'msal.cache.encryption'.
// Electron drops this cookie on process exits, so the encrypted tokens can't be decrypted anymore
// This forces a fresh login on every start (#2681)
// Set an expiration date for the cookie to promote it from a session cookie, so it survives restarts
const MSAL_ENCRYPTION_COOKIE = 'msal.cache.encryption';
function keepMsalEncryptionCookiePersistent(windowSession) {
  // undefined counts as enabled: a partial auth block in config.json replaces
  // the defaults wholesale and must not silently disable this fix (#2722)
  if (config?.auth?.keepMsalCacheEncryptionCookie?.enabled === false) return;
  windowSession.cookies.on('changed', (_event, cookie, _cause, removed) => {
    if (removed || cookie.name !== MSAL_ENCRYPTION_COOKIE || !cookie.session) {
      return;
    }

    const bareDomain = (cookie.domain || '').replace(/^\./, '');
    if (!bareDomain) return;
    const url = `${cookie.secure ? 'https' : 'http'}://${bareDomain}${cookie.path || '/'}`;

    // Get the days to keep the cookie from the config.
    // If no value is set or the value is outside of 1 and 400 or not a number set it to 400
    let keepMsalEncryptionDays = config?.auth?.keepMsalCacheEncryptionCookie?.days ?? 400;
    if(keepMsalEncryptionDays > 400 || keepMsalEncryptionDays < 1 || Number.isNaN(keepMsalEncryptionDays)) keepMsalEncryptionDays = 400;

    // Preserve the original attributes exactly, only add an expiry.
    const details = {
      url,
      name: cookie.name,
      value: cookie.value,
      path: cookie.path,
      secure: cookie.secure,
      httpOnly: cookie.httpOnly,
      sameSite: cookie.sameSite,
      // Keep the Cookie for 400 days
      expirationDate: Math.floor(Date.now() / 1000) + (keepMsalEncryptionDays * 24 * 60 * 60),
    };
    if ((cookie.domain || '').startsWith('.')) {
      details.domain = cookie.domain;
    }

    windowSession.cookies.set(details)
      .then(() => console.debug('[AUTH_RECOVERY] Promoted msal.cache.encryption cookie to persistent (survives restart)'))
      .catch((err) => console.warn('[AUTH_RECOVERY] Failed to persist msal.cache.encryption cookie:', err.message));
  });
}

// Always-on auth-failure signatures. MSAL reports an interaction-required error
// only when a silent token refresh genuinely fails, so it is a reliable signal to
// recover on. It surfaces in two spellings: the MSAL class name
// `InteractionRequired` (console) and the OAuth error code `interaction_required`
// (the lowercase form thrown by token warming as an unhandled rejection — wired
// into detection by the unhandled-rejection handler in app/index.js).
const AUTH_FAILURE_PATTERNS = ['InteractionRequired', 'interaction_required'];
// Opt-in, correlation-only auth-failure signature. A Microsoft 365 web worker
// can die with an uncaught "UPR: <reason>" error, and a genuinely stale session
// emits these (sometimes without ever logging InteractionRequired, #2480). But the worker
// emits the same "UPR:" shape for routine transient failures (pinned channels,
// presence, "Invalid id undefined", and an empty reason) on perfectly healthy
// sessions, and UPR-heavy tenants fire them constantly (#2629) — far too noisy to
// drive an automatic reload. So when auth.reauthRecovery.enabled is set, a UPR
// only records a failure signal for login-popup correlation (so the banner-click
// in-app recovery can recognise a broken session); it never schedules the silent
// clear-and-reload on its own. Recovery from a UPR happens only through an
// explicit user action (the banner click).
const OPT_IN_AUTH_FAILURE_PATTERNS = ['Uncaught Error: UPR:'];
// Only trust auth failure signals from Outlook/Microsoft origins. Outlook's
// scripts are served from the Office CDN, so that counts as a source too.
const TRUSTED_AUTH_SOURCES = [
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.cloud.microsoft',
  'outlook.live.com',
  'cdn.office.net',
  'login.microsoftonline.com',
];
// Loop guard for the automatic clear-and-reload. A cooldown rather than a
// single-shot flag: a long-running app can hit a second stale session hours
// after the first recovery (observed in the field: recovery at 10:55, the
// session went stale again at 12:11 and the old once-per-process flag
// silently swallowed it), while the cooldown still prevents rapid
// reload loops if a recovery fails to fix the session.
let lastAuthRecoveryAt = 0;
const AUTH_RECOVERY_COOLDOWN_MS = 30 * 60 * 1000;
// Timestamp of the last trusted auth-failure signal. Used to correlate login
// popups with an actually-broken session: while the stale "sign in again"
// banner is up the renderer keeps emitting failure signatures every few
// minutes (observed gaps up to ~40 min), whereas healthy-session login popups
// (initial sign-in, consent, step-up MFA, adding an account) appear without
// any preceding failure signal.
let lastAuthFailureSignalAt = 0;
const AUTH_FAILURE_SIGNAL_WINDOW_MS = 60 * 60 * 1000;
// The broken session survives an app restart, so the signal timestamp must
// too (observed in the field: app restarted while stale, the in-memory
// timestamp reset, and the banner's Sign In click fell through unintercepted).
// Persisted via the settings store; restored in onAppReady. Writes are
// throttled because failure signatures can repeat at sub-minute rates.
// Deliberately only feeds popup correlation — automatic recovery stays
// driven by live signals so a stale persisted value can never wipe a
// session that a previous run already fixed.
const AUTH_SIGNAL_STORE_KEY = 'authRecovery.lastFailureSignalAt';
const AUTH_SIGNAL_PERSIST_INTERVAL_MS = 5 * 60 * 1000;
let lastAuthSignalPersistAt = 0;

function recordAuthFailureSignal() {
  lastAuthFailureSignalAt = Date.now();
  if (lastAuthFailureSignalAt - lastAuthSignalPersistAt < AUTH_SIGNAL_PERSIST_INTERVAL_MS) return;
  lastAuthSignalPersistAt = lastAuthFailureSignalAt;
  try {
    appConfig?.settingsStore?.set(AUTH_SIGNAL_STORE_KEY, lastAuthFailureSignalAt);
  } catch (err) {
    console.warn('[AUTH_RECOVERY] Failed to persist failure signal timestamp:', err.message);
  }
}

/**
 * Checks a renderer-originated message (console output or forwarded window
 * error) for auth-failure signatures and schedules recovery on a match.
 * `sourceId` is the script URL the message came from (console-message
 * sourceId or window-error filename).
 */
function maybeScheduleAuthRecovery(message, sourceId) {
  // The whole in-app auth-recovery feature (#2622) is opt-in while it
  // stabilises: with auth.reauthRecovery.enabled off, renderer auth-failure
  // signals are ignored entirely and the app keeps its pre-#2622 behaviour
  // (Outlook's own stale "sign in again" banner stays up; the user relaunches to
  // re-authenticate). The separate #2296 startup/after-sleep cookie cleaning is
  // unaffected and stays always-on.
  if (!config?.auth?.reauthRecovery?.enabled) return;

  const text = message || '';
  // Classify a worker UPR by its shape first, independent of payload contents: a
  // UPR can embed "...code:InteractionRequired..." (StartUpJob/CalendarSyncJob),
  // and matching that as the reliable signal would auto-reload on a UPR despite
  // the correlation-only intent (#2629). So a UPR is never treated as reliable;
  // it only records a correlation signal for the banner interception (the flag is
  // already gated above, so the opt-in worker signal needs no extra flag check).
  const isWorkerUpr = OPT_IN_AUTH_FAILURE_PATTERNS.some(p => text.includes(p));
  const isReliableSignal = !isWorkerUpr && AUTH_FAILURE_PATTERNS.some(p => text.includes(p));
  const isOptInWorkerSignal = isWorkerUpr;
  if (!isReliableSignal && !isOptInWorkerSignal) return;

  // Verify the message originates from a trusted Microsoft source
  const source = sourceId || '';
  if (source && !TRUSTED_AUTH_SOURCES.some(s => source.includes(s))) return;

  // Record the signal even while recovery is cooling down, so the
  // login-popup interception can still correlate against a session that
  // stays broken. For a worker UPR this is the only thing it drives
  // automatically — see the silent-reload guard below.
  recordAuthFailureSignal();

  if (Date.now() - lastAuthRecoveryAt < AUTH_RECOVERY_COOLDOWN_MS) return;

  // Only the reliable interaction-required signal triggers the
  // silent clear-and-reload. A worker UPR is too noisy to auto-recover on
  // (healthy sessions emit them ~hourly, UPR-heavy tenants constantly — #2629);
  // it has already fed popup correlation above, so it can still drive an in-app
  // recovery via the banner-click intercept, just never silently on its own.
  if (!isReliableSignal) return;

  scheduleAuthRecovery();
}

function scheduleAuthRecovery() {
  lastAuthRecoveryAt = Date.now();
  console.info('[AUTH_RECOVERY] Auth failure detected, scheduling recovery');

  // Delay to let Outlook's own retry mechanism attempt recovery first
  setTimeout(
    () =>
      triggerAuthRecovery().catch((err) => {
        console.error('[AUTH_RECOVERY] Failed to trigger auth recovery:', err);
      }),
    5000
  );
}

/**
 * Decides whether an outgoing Microsoft login popup should be intercepted
 * for in-app auth recovery. Requires both the auth.reauthRecovery.enabled
 * opt-in and a trusted auth-failure signal within the correlation window,
 * so login popups from healthy-session flows (initial sign-in, consent,
 * step-up MFA, adding an account) are never diverted — they fall through
 * to the default link handling.
 */
function shouldInterceptAuthPopup() {
  if (!config.auth?.reauthRecovery?.enabled) return false;
  return Date.now() - lastAuthFailureSignalAt < AUTH_FAILURE_SIGNAL_WINDOW_MS;
}

// A single banner click can surface as several requests in quick succession
// (and a broken session's own silent-refresh retries produce more), so
// popup-triggered recovery is deduped over a short window. Unlike the
// automatic path's 30-minute cooldown, a deliberate retry a minute later
// should still work.
let lastPopupRecoveryAt = 0;
const POPUP_RECOVERY_DEDUPE_MS = 10 * 1000;

function triggerPopupRecovery(context) {
  const now = Date.now();
  if (now - lastPopupRecoveryAt < POPUP_RECOVERY_DEDUPE_MS) return;
  lastPopupRecoveryAt = now;
  console.info(`[AUTH_RECOVERY] ${context}, triggering in-app recovery`);
  setImmediate(() =>
    triggerAuthRecovery().catch((err) => {
      console.error('[AUTH_RECOVERY] Failed to trigger auth recovery:', err);
    })
  );
}

/**
 * Clears stale auth state (localStorage tokens + cookies) and reloads
 * the page to force a fresh interactive login.
 */
async function triggerAuthRecovery() {
  console.info('[AUTH_RECOVERY] Clearing auth state and reloading...');

  // Clear localStorage auth tokens via renderer
  try {
    const patternsJson = JSON.stringify(AUTH_LOCAL_STORAGE_PATTERNS);
    const cleared = await window.webContents.executeJavaScript(`
      (function() {
        const patterns = ${patternsJson};
        const keysToRemove = [];
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && patterns.some(p => key.includes(p))) {
            keysToRemove.push(key);
          }
        }
        for (const key of keysToRemove) {
          localStorage.removeItem(key);
        }
        return keysToRemove.length;
      })()
    `);
    console.info('[AUTH_RECOVERY] Cleared localStorage auth entries', { count: cleared });
  } catch (err) {
    console.warn('[AUTH_RECOVERY] Failed to clear localStorage:', err.message);
  }

  await cleanExpiredAuthCookies(window.webContents.session, true);

  console.info('[AUTH_RECOVERY] Reloading for fresh auth...');
  window.loadURL(config.url, { userAgent: config.chromeUserAgent });
}

exports.onAppReady = async function onAppReady(configGroup) {
  appConfig = configGroup;
  config = configGroup.startupConfig;

  const intuneEnabled = config.auth?.intune?.enabled;
  const intuneUser = config.auth?.intune?.user ?? "";
  if (intuneEnabled) {
    intune = require("../intune");
    await intune.initSso(intuneUser);
  }

  if (config.trayIconEnabled) {
    iconChooser = new TrayIconChooser(config);

    if (isMac) {
      console.info("Setting Dock icon for macOS");

      // macOS requires >=128x128 for the dock; use the 256x256 asset by default.
      const DEFAULT_MACOS_DOCK_ICON = "assets/icons/icon-256x256.png";
      const dockIconPath = config.appIcon && config.appIcon.trim() !== ""
        ? config.appIcon
        : path.join(config.appPath, DEFAULT_MACOS_DOCK_ICON);

      const icon = nativeImage.createFromPath(dockIconPath);
      const iconSize = icon.getSize();

      if (iconSize.width < 128) {
        console.warn(
          `Unable to set dock icon for macOS, icon size is less than 128x128, current size ${iconSize.width}x${iconSize.height}. Using resized icon.`
        );
        const resizedIcon = icon.resize({ width: 128, height: 128 });
        app.dock.setIcon(resizedIcon);
      } else {
        app.dock.setIcon(icon);
      }
    }
  }

  const browserWindowManager = new BrowserWindowManager({
    config: config,
    iconChooser: iconChooser,
  });

  window = await browserWindowManager.createWindow();

  connectionManager = new ConnectionManager();

  if (iconChooser) {
    menus = new Menus(window, configGroup, iconChooser.getFile(), connectionManager);
    menus.onSpellCheckerLanguageChanged = onSpellCheckerLanguageChanged;
  }

  addEventHandlers();

  // Keep the msal.cache.encryption cookie
  // Run before loading Outlook so the cookie gets caught right at the start
  keepMsalEncryptionCookiePersistent(window.webContents.session);

  // Clean expired auth cookies before loading Outlook to prevent the
  // "We need you to sign in again" stale banner (#2296)
  await cleanExpiredAuthCookies(window.webContents.session);

  // Restore the last persisted auth-failure signal so login-popup
  // correlation survives app restarts (the broken session does).
  lastAuthFailureSignalAt = Number(appConfig.settingsStore.get(AUTH_SIGNAL_STORE_KEY)) || 0;

  // Monitor renderer auth-failure signals. When Outlook can't refresh tokens
  // silently (e.g., after overnight idle), MSAL logs InteractionRequired. We
  // detect this, clear stale auth state, and reload to force a clean
  // interactive login. Detection itself lives in maybeScheduleAuthRecovery
  // so the forwarded window-error path can reuse it.
  window.webContents.on('console-message', (event) => {
    maybeScheduleAuthRecovery(event.message, event.sourceId);
  });

  login.handleLoginDialogTry(window, config.ssoBasicAuthUser, config.ssoBasicAuthPasswordCommand);

  // A mailto: link that launched the app waits until Outlook has loaded
  // (and the user has signed in) before its compose window opens.
  processArgs(process.argv);
  connectionManager.start(config.url, {
    window: window,
    config: config,
  });

  applyAppConfiguration(config, window);
};

function onSpellCheckerLanguageChanged(languages) {
  appConfig.legacyConfigStore.set("spellCheckerLanguages", languages);
}

let allowFurtherRequests = true;

// Feed forwarded renderer window errors into auth-failure detection. Worker
// uncaught errors (e.g. "Uncaught Error: UPR:") arrive via the window-error
// IPC channel rather than console-message, so app/index.js calls this from
// that handler. `filename` is the originating script URL.
exports.notifyRendererError = function (message, filename) {
  maybeScheduleAuthRecovery(message, filename);
};

exports.show = function () {
  window.show();
};

// Restore if minimised, show if hidden to tray, then focus. Used by the
// notification click handler when notifications.electron.clickAction is
// "restore" (issue #2647).
exports.restoreWindow = restoreWindow;

exports.getWindow = function () {
  return window;
};

exports.onAppSecondInstance = function onAppSecondInstance(event, args) {
  console.debug("second-instance started");
  if (window && !window.isDestroyed()) {
    event.preventDefault();
    // Short cooldown: desktop environments can deliver one click as two
    // launches, but a second mailto: link clicked deliberately must still open.
    if (allowFurtherRequests) {
      allowFurtherRequests = false;
      setTimeout(() => {
        allowFurtherRequests = true;
      }, 1000);
      processArgs(args);
    }

    restoreWindow();
  }
};

// Compose deep link waiting for Outlook to finish loading. Only the newest
// mailto: link is kept.
let pendingComposeUrl = null;
// Set once the main window has settled on an Outlook page, i.e. the user is
// signed in and compose windows can share the session.
let outlookLoaded = false;
// Signed out, Outlook's page finishes loading and only then redirects to the
// Microsoft login page client-side. Treat Outlook as loaded only when the
// window is still on it this long after the load.
const OUTLOOK_SETTLE_MS = 3000;
let outlookSettleTimer = null;

/**
 * Opens an Outlook compose deep link in its own window, or queues it until
 * Outlook has loaded in the main window. The compose window shares the main
 * window's session partition, so it is already signed in.
 *
 * @param {string} composeUrl - Outlook compose deep link
 */
function openComposeWindow(composeUrl) {
  if (!outlookLoaded) {
    pendingComposeUrl = composeUrl;
    return;
  }
  pendingComposeUrl = null;

  const composeWindow = new BrowserWindow({
    width: 1000,
    height: 760,
    autoHideMenuBar: true,
    icon: iconChooser ? nativeImage.createFromPath(iconChooser.getFile()) : undefined,
    webPreferences: {
      partition: config.partition,
      contextIsolation: true,
      sandbox: true,
      spellcheck: true,
    },
  });
  composeWindow.webContents.setWindowOpenHandler(onComposeWindowOpen);
  // Outlook calls window.close() once the message is sent or discarded.
  composeWindow.loadURL(composeUrl, { userAgent: config.chromeUserAgent }).catch(() => {
    // The URL carries recipients and subject; never log it.
    console.debug("[MAILTO] compose window navigation failed");
  });
}

// Links clicked inside a compose window (signatures, help) open externally,
// except Outlook's own pages, which stay in-app.
function onComposeWindowOpen(details) {
  if (isOutlookUrl(details.url)) {
    return { action: "allow" };
  }
  if (/^https?:/i.test(details.url)) {
    openInBrowser(details);
  }
  return { action: "deny" };
}

function releasePendingCompose() {
  if (pendingComposeUrl) {
    console.info("[MAILTO] Outlook loaded, opening the queued compose window");
    openComposeWindow(pendingComposeUrl);
  }
}

function applyAppConfiguration(config, window) {
  applySpellCheckerConfiguration(config.spellCheckerLanguages, window);

  const certPath = config.clientCertPath;
  if (certPath) {
    app.importCertificate(
      {
        certificate: certPath,
        password: config.clientCertPassword || "",
      },
      (result) => {
        console.info(
          `[CERT] Client certificate loaded, result: ${result}`
        );
      }
    );
  }
  window.webContents.setUserAgent(config.chromeUserAgent);

  if (!config.minimized) {
    window.show();
  } else {
    window.hide();
  }

  if (config.webDebug) {
    window.openDevTools();
  }
}

function applySpellCheckerConfiguration(languages, window) {
  const spellCheckProvider = new SpellCheckProvider(window);
  if (
    spellCheckProvider.setLanguages(languages).length === 0 &&
    languages.length > 0
  ) {
    // If failed to set user supplied languages, fallback to system locale.
    const systemList = [app.getLocale()];
    if (app.getLocale() !== app.getSystemLocale()) {
      systemList.push(app.getSystemLocale());
    }
    spellCheckProvider.setLanguages(systemList);
  }
}

function onDidFinishLoad() {
  console.debug("did-finish-load");

  // Skip CSS injection on Chrome error pages (e.g. chrome-error://chromewebdata/)
  // which appear when navigation fails due to network errors like ERR_NAME_NOT_RESOLVED.
  const currentUrl = window.webContents.getURL();
  if (!currentUrl.startsWith("https://")) {
    console.debug(`[CONNECTION] Skipping page setup on non-Outlook page: ${currentUrl.split("?")[0]}`);
    return;
  }

  customCSS.onDidFinishLoad(window.webContents, config);

  if (isOutlookUrl(currentUrl)) {
    clearTimeout(outlookSettleTimer);
    outlookSettleTimer = setTimeout(() => {
      if (window && !window.isDestroyed() && isOutlookUrl(window.webContents.getURL())) {
        outlookLoaded = true;
        releasePendingCompose();
      }
    }, OUTLOOK_SETTLE_MS);
  }
}

// Leaving Outlook (sign-out, session expiry, auth recovery) means compose
// windows would land on the login page, so queue mailto: links again.
function onMainFrameNavigated(_event, url) {
  if (!isOutlookUrl(url)) {
    outlookLoaded = false;
  }
}

function restoreWindow() {
  if (window.isMinimized()) {
    window.restore();
  } else if (!window.isVisible()) {
    window.show();
  }

  window.focus();
  // A second instance has no user activation, so window.focus() alone can be
  // refused as focus stealing; app.focus() raises the first visible window on
  // X11 (Wayland may only flash the icon, which is the platform's call).
  app.focus();
}

/**
 * Processes command line arguments for a mailto: link and opens (or queues)
 * the matching Outlook compose window.
 *
 * @param {string[]} args - Command line arguments to process
 */
function processArgs(args) {
  if (config.mailto?.enabled === false) return;

  const mailtoArg = mailto.findMailtoArg(args);
  if (!mailtoArg) return;

  const composeUrl = mailto.toComposeUrl(mailtoArg, config.url);
  if (!composeUrl) {
    console.warn("[MAILTO] Ignoring a mailto: link that could not be translated");
    return;
  }
  console.info("[MAILTO] mailto: link received");
  openComposeWindow(composeUrl);
}

// Microsoft telemetry / beacon hosts that are not required for Outlook to
// function. Blocking these at webRequest cancels both the network traffic
// and the downstream sub-frame failure logs they would otherwise produce
// in restricted-network environments. Kept deliberately narrow: anything
// Outlook needs to function (outlook.office.com, *.office.net,
// login.microsoftonline.com, *.trafficmanager.net) is excluded. Start
// with this initial set and expand as new hosts are confirmed safe to
// drop; any new entry must also satisfy `MS_TELEMETRY_FAST_PATH` below
// or the fast-path string must be updated.
const MS_TELEMETRY_HOSTS = [
  'events.data.microsoft.com',
  'browser.events.data.msn.com',
];

// Substring guard cheap-checked before the URL parse below. Every entry
// in `MS_TELEMETRY_HOSTS` must contain this substring so the fast path
// never produces a false negative.
const MS_TELEMETRY_FAST_PATH = 'events.data.';

function isMicrosoftTelemetryHost(url) {
  // Fast path: avoid `new URL(...)` on every HTTPS request. The handler
  // fires for every request matched by `{ urls: ["https://*/*"] }`, so
  // skipping the parse for the overwhelmingly common non-telemetry case
  // is measurable on chat-heavy sessions.
  if (!url?.includes(MS_TELEMETRY_FAST_PATH)) return false;
  try {
    const hostname = new URL(url).hostname;
    return MS_TELEMETRY_HOSTS.some(
      (h) => hostname === h || hostname.endsWith('.' + h)
    );
  } catch {
    return false;
  }
}

function onBeforeRequestHandler(details, callback) {
  if (isMicrosoftTelemetryHost(details.url)) {
    callback({ cancel: true });
    return;
  }

  if (aboutBlankRequestCount < 1) {
    callback({});
  } else if (details.resourceType === "mainFrame") {
    // A top-level navigation is never the about:blank popup's own request, so
    // it must not be diverted into the hidden child window below. Diverting it
    // cancels the navigation (ERR_BLOCKED_BY_CLIENT) and leaves a blank page,
    // e.g. the guest / number-matching MFA sign-in where the main frame
    // navigates to the authorize URL right after an about:blank popup bumped
    // the counter (#2591). A new top-level navigation also makes any pending
    // interceptions stale, so reset the counter to 0 rather than decrementing
    // it: that way a leftover count cannot divert the new page's sub-resources.
    aboutBlankRequestCount = 0;
    callback({});
  } else if (isAuthLoginUrl(details.url) && shouldInterceptAuthPopup()) {
    // The hidden-window handling below exists for SILENT token refresh. An
    // interactive login can never complete in a window that is hidden and
    // destroyed on ready-to-show — which is why clicking the stale "sign in
    // again" banner appears to do nothing: the click opens an about:blank
    // popup whose login navigation lands here and dies invisibly (it never
    // reaches the isAuthLoginUrl check in onNewWindow, because the URL is
    // still about:blank at window-open time). When the navigation correlates
    // with a broken session (see shouldInterceptAuthPopup), run in-app
    // recovery instead. Reset the counter: recovery reloads the page, so any
    // pending interceptions are stale (#2591 rationale).
    aboutBlankRequestCount = 0;
    triggerPopupRecovery('Login navigation from about:blank popup intercepted');
    callback({ cancel: true });
  } else {
    // Open request in hidden child window for authentication
    const child = new BrowserWindow({ parent: window, show: false });
    child.loadURL(details.url);
    child.once("ready-to-show", () => {
      child.destroy();
    });

    aboutBlankRequestCount -= 1;
    callback({ cancel: true });
  }
}

// Microsoft Identity Platform login hostnames. When Outlook opens a popup to
// one of these it is requesting interactive re-authentication (e.g. the
// "sign in again" banner). Kept separate from AUTH_DOMAINS because that
// list includes broad domains used for cookie scoping; this narrower set
// is only the endpoints that initiate an OAuth/OIDC interactive flow.
const AUTH_LOGIN_DOMAINS = [
  'login.microsoftonline.com',
  'login.microsoft.com',
  'login.live.com',
];

/**
 * Returns true when the URL targets a Microsoft Identity Platform login page.
 * Used to intercept re-auth popups that Outlook opens from the "sign in again"
 * banner so they complete inside the Electron app instead of opening an
 * external browser window that Electron cannot observe.
 */
function isAuthLoginUrl(url) {
  try {
    const hostname = stripMcasSuffix(new URL(url).hostname);
    return AUTH_LOGIN_DOMAINS.some(d => hostname === d || hostname.endsWith('.' + d));
  } catch {
    return false;
  }
}

/**
 * Strips report-only CSP headers for non-Outlook domains (#2326).
 *
 * With contextIsolation disabled the shared V8 context erroneously
 * enforces report-only policies as blocking, breaking SSO flows
 * that rely on dynamic code or nonce-less scripts (e.g. Symantec VIP).
 * Report-only headers are safe to strip since they should never block.
 */
function stripCspForAuthPages(responseHeaders, url) {
  try {
    if (isOutlookHost(stripMcasSuffix(new URL(url).hostname))) return;
  } catch {
    return;
  }

  for (const key of Object.keys(responseHeaders)) {
    if (key.toLowerCase() === 'content-security-policy-report-only') {
      let hostname;
      try {
        hostname = new URL(url).hostname;
      } catch {
        hostname = 'unknown';
      }
      console.debug(`[CSP] Stripping report-only header from: ${hostname}`);
      delete responseHeaders[key];
    }
  }
}

function onHeadersReceivedHandler(details, callback) {
  stripCspForAuthPages(details.responseHeaders, details.url);

  callback({
    responseHeaders: details.responseHeaders,
  });
}

function onBeforeSendHeadersHandler(detail, callback) {
  if (intune?.isSsoUrl(detail.url)) {
    intune.addSsoCookie(detail, callback);
  } else {
    callback({
      requestHeaders: detail.requestHeaders,
    });
  }
}

function onNewWindow(details) {
  // Breadcrumb for diagnosing which shape of popup the "sign in again"
  // banner opens (about:blank-then-navigate vs a direct login URL) — the two
  // take different paths below and only the direct form can be intercepted
  // for recovery. Auth-related popups only, origin only: no path, query, or
  // user-specific data is logged.
  if (details.url.startsWith("about:blank") || isAuthLoginUrl(details.url)) {
    let origin = "about:blank";
    if (!details.url.startsWith("about:blank")) {
      try {
        origin = new URL(details.url).origin;
      } catch {
        origin = "unparseable";
      }
    }
    console.info('[WINDOW_OPEN] Auth-related popup', { origin });
  }

  if (
    details.url === "about:blank" ||
    details.url === "about:blank#blocked"
  ) {
    aboutBlankRequestCount += 1;
    return { action: "deny" };
  } else if (isAuthLoginUrl(details.url) && shouldInterceptAuthPopup()) {
    // Outlook is opening a direct-URL Microsoft login popup from a session
    // that recently emitted auth-failure signals (see
    // shouldInterceptAuthPopup). Opening login.microsoftonline.com in an
    // external browser completes auth there but Electron never receives the
    // result, so trigger in-app recovery instead: clear stale auth state and
    // reload Outlook for a fresh interactive sign-in within the app. (The
    // stale-banner popup is usually about:blank-shaped and is handled in
    // onBeforeRequestHandler; this branch covers the direct-URL form.)
    // Login popups without a correlated failure signal (initial sign-in,
    // consent and step-up prompts, adding an account) keep the original
    // open-externally behaviour via secureOpenLink below.
    triggerPopupRecovery('Direct login popup intercepted');
    return { action: "deny" };
  } else if (isOutlookUrl(details.url)) {
    // Outlook's own pop-outs (a message or draft opened in a new window, a
    // calendar event) stay in-app and signed in. The child inherits the
    // main window's preload, so it is flagged as a pop-out for the preload
    // to skip the tray and unread-count tools.
    return {
      action: "allow",
      overrideBrowserWindowOptions: {
        autoHideMenuBar: true,
        webPreferences: {
          additionalArguments: [POPOUT_WINDOW_ARG],
        },
      },
    };
  }

  return secureOpenLink(details);
}

function onNavigationChanged() {
  if (window?.webContents?.navigationHistory) {
    const canGoBack = window.webContents.navigationHistory.canGoBack();
    const canGoForward = window.webContents.navigationHistory.canGoForward();
    window.webContents.send("navigation-state-changed", canGoBack, canGoForward);
  }
}

function onWindowClosed() {
  console.debug("window closed");
  window = null;
  app.quit();
}

function addEventHandlers() {
  // After resuming from sleep, check if auth cookies expired during suspend.
  // Electron on Linux lacks OS-level auth brokers (WAM/Keychain) that browsers
  // use to transparently refresh tokens, so we handle expiry ourselves.
  const { powerMonitor } = require("electron");
  powerMonitor.on("resume", async () => {
    console.debug('[AUTH_RECOVERY] System resumed, checking auth cookies');
    const result = await cleanExpiredAuthCookies(window.webContents.session);
    if (result.expired > 0) {
      console.info('[AUTH_RECOVERY] Cleaned expired cookies after resume', {
        cleaned: result.cleaned,
        expired: result.expired,
      });
      // Let Outlook's own MSAL retry handle re-authentication rather than
      // triggering full recovery which clears all auth state (issue #2364)
    }
  });

  window.webContents.setWindowOpenHandler(onNewWindow);
  // Pop-outs allowed by onNewWindow get the same policy, so a link in a
  // message opened in its own window still goes to the system browser.
  window.webContents.on("did-create-window", (child) => {
    child.webContents.setWindowOpenHandler(onNewWindow);
  });
  window.webContents.session.webRequest.onBeforeRequest(
    { urls: ["https://*/*"] },
    onBeforeRequestHandler
  );
  window.webContents.session.webRequest.onHeadersReceived(
    { urls: ["https://*/*"] },
    onHeadersReceivedHandler
  );
  window.webContents.session.webRequest.onBeforeSendHeaders(
    getWebRequestFilterFromURL(),
    onBeforeSendHeadersHandler
  );
  window.webContents.on("did-finish-load", onDidFinishLoad);
  window.on("closed", onWindowClosed);
  window.webContents.addListener("before-input-event", onBeforeInput);

  // Navigation state change handlers
  window.webContents.on("did-navigate", onNavigationChanged);
  window.webContents.on("did-navigate", onMainFrameNavigated);
  window.webContents.on("did-navigate-in-page", onNavigationChanged);

  // Pre-fill/advance the Microsoft/federated web login page (no-op unless one
  // of auth.webLogin.user / auth.webLogin.passwordCommand / auth.webLogin.verifyMethod is set).
  ssoPasswordPrefill.attach(window, config);
}

function getWebRequestFilterFromURL() {
  const filter = { urls: ["https://*/*"] };
  if (intune) {
    intune.setupUrlFilter(filter);
  }

  return filter;
}

function onBeforeInput(event, input) {
  isControlPressed = input.control;

  if (input.type !== "keyDown") {
    return;
  }
  const history = window?.webContents?.navigationHistory;
  if (!history) {
    return;
  }

  // Keyboard history navigation, independent of the Outlook DOM. Keys are
  // platform-specific: on macOS,
  // Option(Alt)+Left/Right is the system word-navigation shortcut inside text
  // fields, so stealing it would break message editing — macOS uses the
  // standard Cmd+[ / Cmd+] instead, while other platforms use the
  // browser-standard Alt+Left / Alt+Right.
  const isMac = process.platform === "darwin";
  const modifierActive = isMac
    ? input.meta && !input.control && !input.alt && !input.shift
    : input.alt && !input.control && !input.meta && !input.shift;
  if (!modifierActive) {
    return;
  }

  const backKey = isMac ? "[" : "ArrowLeft";
  const forwardKey = isMac ? "]" : "ArrowRight";
  if (input.key === backKey && history.canGoBack()) {
    event.preventDefault();
    history.goBack();
  } else if (input.key === forwardKey && history.canGoForward()) {
    event.preventDefault();
    history.goForward();
  }
}

function secureOpenLink(details) {
  console.debug('[LINK] Requesting to open external link');
  const action = getLinkAction();

  if (action === 0) {
    openInBrowser(details);
  }

  const returnValue =
    action === 1
      ? {
          action: "allow",
          overrideBrowserWindowOptions: {
            modal: true,
            useContentSize: true,
            parent: window,
          },
        }
      : { action: "deny" };

  if (action === 1) {
    removePopupWindowMenu();
  }

  return returnValue;
}

function openInBrowser(details) {
  if (config.defaultURLHandler.trim() === "") {
    shell.openExternal(details.url);
  } else {
    execFile(
      config.defaultURLHandler.trim(),
      [details.url],
      openInBrowserErrorHandler
    );
  }
}

function openInBrowserErrorHandler(error) {
  if (error) {
    console.error(`openInBrowserErrorHandler ${error.message}`);
  }
}

function getLinkAction() {
  const action = isControlPressed
    ? dialog.showMessageBoxSync(window, {
        type: "warning",
        buttons: ["Allow", "Deny"],
        title: "Open URL",
        normalizeAccessKeys: true,
        defaultId: 1,
        cancelId: 1,
        message:
          "This will open the URL in the application context. If this is for SSO, click Allow otherwise Deny.",
      }) + 1
    : 0;

  isControlPressed = false;
  return action;
}

async function removePopupWindowMenu() {
  for (let i = 1; i <= 200; i++) {
    await sleep(10);
    const childWindows = window.getChildWindows();
    if (childWindows.length) {
      childWindows[0].removeMenu();
      break;
    }
  }
}

async function sleep(ms) {
  return await new Promise((r) => setTimeout(r, ms));
}
