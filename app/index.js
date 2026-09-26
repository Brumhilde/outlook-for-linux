const {
  app,
  dialog,
  ipcMain,
  globalShortcut,
  nativeImage,
  nativeTheme,
  session,
} = require("electron");
const path = require("node:path");
const crypto = require("node:crypto");
const { allowedChannels } = require("./security/ipcValidator");
const { installIpcSecurity } = require("./security/ipcSecurity");
const { sanitize: sanitizePii } = require("./utils/logSanitizer");
const { register: registerGlobalShortcuts } = require("./globalShortcuts");
const CommandLineManager = require("./startup/commandLine");
const NotificationService = require("./notifications/service");
const CustomNotificationManager = require("./notificationSystem");
const DownloadManager = require("./downloadManager");
const PartitionsManager = require("./partitions/manager");
const AutoUpdater = require("./autoUpdater");
const WebAuthn = require("./webauthn");
const mailto = require("./mailto");
const os = require("node:os");

// Name notifications "Outlook for Linux" instead of the raw app name.
// Desktop environments read two different fields for the header: KDE shows
// the notification's app_name (which Electron takes from app.name), GNOME
// resolves the desktop-entry hint and shows that entry's Name. Cover both.
// app.name also seeds the default userData path, so pin the path first or
// existing installs would silently switch config directory.
if (process.platform === "linux") {
  const userDataPath = app.getPath("userData");
  app.setName("Outlook for Linux");
  app.setPath("userData", userDataPath);

  // The desktop-entry hint must be set unconditionally: Electron pre-sets
  // CHROME_DESKTOP itself before app code runs, so guarding on the env var
  // being absent would be a no-op. Sandboxed installs export the entry
  // under their own id.
  if (process.env.FLATPAK_ID) {
    app.setDesktopName(`${process.env.FLATPAK_ID}.desktop`);
  } else if (process.env.SNAP_INSTANCE_NAME) {
    app.setDesktopName(`${process.env.SNAP_INSTANCE_NAME}_outlook-for-linux.desktop`);
  } else {
    app.setDesktopName("outlook-for-linux.desktop");
  }
}

const { NETWORK_ERROR_PATTERNS } = require("./config/defaults");

function isNetworkError(message) {
  if (typeof message !== 'string') return false;
  if (NETWORK_ERROR_PATTERNS.some(pattern => message.includes(pattern))) return true;
  // "Object has been destroyed" errors can occur when the window is destroyed
  // during network-triggered operations (e.g., reload after network recovery).
  // These are transient and should not terminate the app.
  if (message.includes('Object has been destroyed')) return true;
  // "Script failed to execute" occurs when executeJavaScript runs on a page where
  // APIs are unavailable (e.g., Chrome error pages after ERR_NAME_NOT_RESOLVED).
  // This is a symptom of network failure, not a fatal error.
  if (message.includes('Script failed to execute')) return true;
  return false;
}

// Top-level error handlers for crash diagnostics
process.on('uncaughtException', (error) => {
  const message = error instanceof Error ? error.message : String(error);
  const stack = error instanceof Error ? error.stack : undefined;
  if (isNetworkError(message)) {
    console.error('[ERROR] Network-related uncaught exception (not terminating):', { message });
    return;
  }
  console.error('[FATAL] Uncaught exception:', { message, stack });
  process.exit(1);
});

process.on('unhandledRejection', (reason) => {
  const message = reason instanceof Error ? reason.message : String(reason);
  const stack = reason instanceof Error ? reason.stack : undefined;
  if (isNetworkError(message)) {
    console.error('[ERROR] Network-related unhandled rejection (not terminating):', { message });
    return;
  }
  console.error('[FATAL] Unhandled promise rejection:', { message, stack });
  process.exit(1);
});

// Support for E2E testing: use temporary userData directory for clean state
if (process.env.E2E_USER_DATA_DIR) {
  app.setPath("userData", process.env.E2E_USER_DATA_DIR);
}

// This must be executed before loading the config file.
CommandLineManager.addSwitchesBeforeConfigLoad();

const { AppConfiguration } = require("./appConfiguration");
const appConfig = new AppConfiguration(
  app.getPath("userData"),
  app.getVersion()
);

const config = appConfig.startupConfig;
config.appPath = path.join(__dirname, app.isPackaged ? "../../" : "");

CommandLineManager.addSwitchesAfterConfigLoad(config);

const { createPlayer } = require("./audio/player");
const player = createPlayer();

const certificateModule = require("./certificate");
const clientCertificate = require("./clientCertificate");
const backgroundPortal = require("./backgroundPortal");
const CacheManager = require("./cacheManager");
const gotTheLock = app.requestSingleInstanceLock();
const mainAppWindow = require("./mainAppWindow");

const notificationService = new NotificationService(
  player,
  config,
  mainAppWindow
);

const partitionsManager = new PartitionsManager(appConfig.settingsStore);

const customNotificationManager = new CustomNotificationManager(config, mainAppWindow);

// Issue #2512: Electron silently saves downloads to ~/Downloads with no UI
// feedback unless `session.on('will-download', …)` is wired up. The manager
// itself attaches to the Outlook session inside handleAppReady once the main
// window has been created (the partition is provisioned at that point).
// `mainAppWindow` is passed so the manager can drive the taskbar progress bar.
// On Linux, wire in two D-Bus emitters that complement each other:
//   - `jobViewEmitter`: per-download progress in KDE Plasma's notification
//     widget (same surface users see for Firefox / Dolphin / KIO).
//   - `launcherEntryEmitter`: aggregate progress on the dock icon for
//     GNOME/Ubuntu users running Ubuntu Dock or Dash-to-Dock (the largest
//     Linux DE audience). Bypasses Electron's `setProgressBar` Linux gate.
// Both degrade to no-op if their respective service isn't on the bus, so
// non-KDE / non-GNOME setups fall back to the window-title `[N%]` prefix.
let jobEmitter = null;
let launcherEmitter = null;
if (os.platform() === "linux") {
  try {
    jobEmitter = require("./downloadManager/jobViewEmitter");
  } catch (error) {
    console.warn("[DownloadManager] JobView emitter unavailable", {
      message: error.message,
    });
  }
  try {
    launcherEmitter = require("./downloadManager/launcherEntryEmitter");
  } catch (error) {
    console.warn("[DownloadManager] LauncherEntry emitter unavailable", {
      message: error.message,
    });
  }
}
const downloadManager = new DownloadManager(config, mainAppWindow, jobEmitter, launcherEmitter);

// Offer the app as the system mailto: handler. Desktop environments still let
// the user pick the default; on Linux the packaged .desktop file also declares
// x-scheme-handler/mailto (see package.json build.protocols).
if (config.mailto?.enabled !== false && !app.isDefaultProtocolClient(mailto.PROTOCOL, process.execPath)) {
  app.setAsDefaultProtocolClient(mailto.PROTOCOL, process.execPath);
}

if (gotTheLock) {
  app.on("second-instance", mainAppWindow.onAppSecondInstance);
  app.on("ready", handleAppReady);
  app.on("quit", () => console.debug("quit"));
  app.on("render-process-gone", onRenderProcessGone);
  app.on("will-quit", () => console.debug("will-quit"));
  app.on("certificate-error", handleCertificateError);
  app.on("browser-window-focus", handleGlobalShortcutDisabled);
  app.on("browser-window-blur", handleGlobalShortcutDisabledRevert);

  // IPC Security: wrap handler-registration methods so every renderer-initiated
  // IPC call is validated against the allowlist in app/security/ipcValidator.js.
  // The wrapping also covers removal, so listeners registered per short-lived
  // window can actually be taken off again. See app/security/ipcSecurity.js.
  installIpcSecurity(ipcMain);

  // Restart application when configuration file changes
  ipcMain.on("config-file-changed", restartApp);
  // Get current application configuration
  ipcMain.handle("get-config", async () => {
    return config;
  });

  notificationService.initialize();
  partitionsManager.initialize();
  customNotificationManager.initialize();

  // Set application badge count (dock/taskbar notification)
  ipcMain.handle("set-badge-count", setBadgeCountHandler);

  // Get application version number
  ipcMain.handle("get-app-version", async () => {
    return config.appVersion;
  });

  // Navigate back in browser history
  ipcMain.on("navigate-back", (event) => {
    const webContents = event.sender;
    if (webContents?.navigationHistory?.canGoBack()) {
      console.debug("Navigating back");
      webContents.navigationHistory.goBack();
    }
  });

  // Navigate forward in browser history
  ipcMain.on("navigate-forward", (event) => {
    const webContents = event.sender;
    if (webContents?.navigationHistory?.canGoForward()) {
      console.debug("Navigating forward");
      webContents.navigationHistory.goForward();
    }
  });

  // Get current navigation state (can go back/forward)
  ipcMain.handle("get-navigation-state", (event) => {
    const webContents = event.sender;
    return {
      canGoBack: webContents?.navigationHistory?.canGoBack() || false,
      canGoForward: webContents?.navigationHistory?.canGoForward() || false,
    };
  });

  // Log renderer-side unhandled promise rejections
  ipcMain.on("unhandled-rejection", (_event, errorData) => {
    // Payload is constructed and length-capped in app/browser/preload.js;
    // prior to this handler + the ipcValidator allowlist entry these
    // messages were silently dropped. Fields are run through
    // sanitizeRendererLogField to scrub PII before logging.
    try {
      // Run the noise check on the raw message — the sanitizer may strip
      // query strings / fragments that contain the auth error code.
      const preLoginNoise = isPreLoginAuthNoise(errorData?.message);
      const log = preLoginNoise ? console.debug : console.error;
      log("[Renderer] Unhandled rejection:", {
        message: sanitizeRendererLogField(errorData?.message, "unknown"),
        stack: sanitizeRendererLogField(errorData?.stack),
        timestamp: toFiniteNumber(errorData?.timestamp, Date.now()),
      });
      // Some auth failures only surface as unhandled promise rejections from MSAL
      // token warming (e.g. "interaction_required"/"InteractionRequired" from
      // acquireTokenV2) and never hit console-message or window-error — feed the
      // raw (unsanitized) message to auth-failure detection. Rejections carry no
      // source URL, so detection's trusted-source check is skipped (empty source).
      // The pre-login-noise check above only down-levels the LOG; it must not
      // gate detection: the reliable "InteractionRequired" signal arrives inside
      // these noise-matched messages, so always forward. maybeScheduleAuthRecovery
      // does its own filtering and only acts on InteractionRequired /
      // interaction_required (login_required and AuthFailed are logged, not acted on).
      mainAppWindow.notifyRendererError(errorData?.message, undefined);
    } catch (err) {
      console.error("[Renderer] Failed to log unhandled-rejection:", err);
    }
  });

  // Log renderer-side uncaught window errors
  ipcMain.on("window-error", (_event, errorData) => {
    try {
      const preLoginNoise = isPreLoginAuthNoise(errorData?.message);
      const log = preLoginNoise ? console.debug : console.error;
      log("[Renderer] Window error:", {
        message: sanitizeRendererLogField(errorData?.message, "unknown"),
        filename: sanitizeRendererLogField(errorData?.filename, "") || "",
        lineno: toFiniteNumber(errorData?.lineno, 0),
        colno: toFiniteNumber(errorData?.colno, 0),
        errorStack: sanitizeRendererLogField(errorData?.errorStack),
        timestamp: toFiniteNumber(errorData?.timestamp, Date.now()),
      });
      // Some auth failures only surface as uncaught worker errors (e.g.
      // "Uncaught Error: UPR:") that never hit the console-message path —
      // feed the raw (unsanitized) message to auth-failure detection. As with
      // the unhandled-rejection handler, pre-login-noise down-levelling controls
      // only the log level, never whether detection sees the signal.
      mainAppWindow.notifyRendererError(errorData?.message, errorData?.filename);
    } catch (err) {
      console.error("[Renderer] Failed to log window-error:", err);
    }
  });
} else {
  console.info("App already running");
  app.quit();
}

function restartApp() {
  console.info("Restarting app...");
  app.relaunch();
  app.exit();
}

const MAX_RENDERER_LOG_FIELD_LENGTH = 4096;

/**
 * Sanitizes a renderer-supplied log string before it hits main-process logs.
 * Runs it through the shared PII sanitizer (which scrubs emails, UUIDs,
 * tokens, IPs, URL query strings, user paths, etc.) and length-caps the
 * result. Non-strings fall back to the supplied fallback.
 *
 * Renderer errors from Outlook can contain UserId/tenantId/Trace ID/Correlation
 * ID values that match the shared UUID pattern. CLAUDE.md treats account IDs
 * and SSO info as must-not-log.
 *
 * @param {unknown} value - The field value to sanitize.
 * @param {any} fallback - Value to return when `value` is not a string.
 * @returns {string|any} - Sanitized string, or `fallback` for non-strings.
 */
function sanitizeRendererLogField(value, fallback = null) {
  if (typeof value !== "string") return fallback;

  // logSanitizer scrubs URL query strings but not fragments. OAuth implicit-flow
  // tokens land in the fragment (#access_token=…), so strip those first to
  // preserve the previous behaviour before running the shared sanitizer.
  const fragmentStripped = value.replaceAll(
    /(\b[a-z][a-z0-9+.-]*:\/\/[^\s#)'"<>]+)#[^\s)'"<>]*/gi,
    "$1#[redacted]",
  );

  const sanitized = sanitizePii(fragmentStripped);

  return sanitized.length > MAX_RENDERER_LOG_FIELD_LENGTH
    ? `${sanitized.slice(0, MAX_RENDERER_LOG_FIELD_LENGTH)}…`
    : sanitized;
}

/**
 * Coerces a renderer-supplied scalar to a finite number, falling back to a
 * caller-supplied default for anything non-numeric or NaN/Infinity. Prevents
 * arbitrary objects or strings from leaking into main-process logs via the
 * error-forwarding IPC channels.
 *
 * @param {unknown} value - The field value to coerce.
 * @param {number} fallback - Value returned when `value` isn't a finite number.
 * @returns {number}
 */
function toFiniteNumber(value, fallback = 0) {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

// The Microsoft 365 web apps flood the renderer with these error signatures whenever no user is
// signed in, until the auth-recovery layer (see app/mainAppWindow/index.js)
// clears stale state and reloads. The recovery layer is the actionable
// channel for these; mirroring them as console.error in our logs just buries
// real issues. Down-leveling to console.debug keeps them available without
// the noise.
const PRE_LOGIN_AUTH_NOISE_PATTERNS = [
  "login_required",
  "aadsts50058",
  "interactionrequired",
  "authfailed",
];

function isPreLoginAuthNoise(message) {
  if (typeof message !== "string") return false;
  const lower = message.toLowerCase();
  return PRE_LOGIN_AUTH_NOISE_PATTERNS.some((p) => lower.includes(p));
}

// A gone renderer (the Outlook PWA) is usually unrecoverable — continuing
// leaves a blank or unresponsive window, so quit for a clean restart.
function onRenderProcessGone(event, webContents, details) {
  console.error(`render-process-gone ${JSON.stringify(details)}`);
  app.quit();
}

function onAppTerminated() {
  app.quit();
}

// Content-based hash so any reworded or new warning re-surfaces — users can
// only dismiss the exact text they have read.
const ACK_WARNINGS_KEY = "warnings.acknowledged";
function hashWarning(warning) {
  return crypto
    .createHash("sha256")
    .update(warning)
    .digest("hex")
    .slice(0, 16);
}

async function showConfigurationDialogs() {
  if (config.error) {
    await dialog.showMessageBox({
      title: "Configuration Error",
      icon: nativeImage.createFromPath(
        path.join(config.appPath, "assets/icons/setting-error.256x256.png")
      ),
      message: `Error in config file '${config.error}'.\n Loading default configuration`,
    });
  }

  if (!config.warnings || config.warnings.length === 0) return;

  const acknowledged = new Set(
    appConfig.settingsStore.get(ACK_WARNINGS_KEY, [])
  );
  // Hash once per warning and dedupe by hash — identical text in
  // `config.warnings` would otherwise prompt twice.
  const pending = new Map();
  for (const warning of config.warnings) {
    const hash = hashWarning(warning);
    if (acknowledged.has(hash) || pending.has(hash)) continue;
    pending.set(hash, warning);
  }
  if (pending.size === 0) return;

  const icon = nativeImage.createFromPath(
    path.join(config.appPath, "assets/icons/alert-diamond.256x256.png")
  );

  let dirty = false;
  for (const [hash, warning] of pending) {
    const { checkboxChecked } = await dialog.showMessageBox({
      title: "Configuration Warning",
      icon,
      message: warning,
      checkboxLabel: "Don't show this again",
    });
    if (checkboxChecked) {
      acknowledged.add(hash);
      dirty = true;
    }
  }

  if (dirty) {
    appConfig.settingsStore.set(ACK_WARNINGS_KEY, Array.from(acknowledged));
  }
}

function loadMenuToggleSettings() {
  const menuToggleSettings = [
    'disableNotifications',
    'disableNotificationSound',
    'disableNotificationSoundIfNotAvailable',
    'disableNotificationWindowFlash',
    'disableBadgeCount',
    'defaultNotificationUrgency',
    'appIcon'
  ];

  for (const setting of menuToggleSettings) {
    if (appConfig.legacyConfigStore.has(setting)) {
      config[setting] = appConfig.legacyConfigStore.get(setting);
    }
  }
}

function initializeCacheManagement() {
  if (!config.cacheManagement?.enabled) return;

  const cacheManager = new CacheManager({
    maxCacheSizeMB: config.cacheManagement?.maxCacheSizeMB || 600,
    cacheCheckIntervalMs:
      config.cacheManagement?.cacheCheckIntervalMs || 60 * 60 * 1000,
    partition: config.partition,
  });
  cacheManager.start();

  app.on("before-quit", () => {
    cacheManager.stop();
  });
}

function initializeAutoUpdater() {
  const mainWindow = mainAppWindow.getWindow();
  if (mainWindow) {
    AutoUpdater.initialize(mainWindow);
  }
}

async function handleAppReady() {
  try {
    await showConfigurationDialogs();

    process.on("SIGTRAP", onAppTerminated);
    process.on("SIGINT", onAppTerminated);
    process.on("SIGTERM", onAppTerminated);
    process.stdout.on("error", () => {});

    initializeCacheManagement();

    loadMenuToggleSettings();

    // Outlook renders its own light/dark theme from prefers-color-scheme, so
    // following the desktop theme only needs Chromium to report it.
    if (config.followSystemTheme) {
      nativeTheme.themeSource = "system";
    }

    // Smartcard / NSS client-certificate PIN dialog (Linux only, issue #2639).
    // Registered before the main window loads so the handler exists before the
    // first TLS handshake that may need a token PIN.
    if (process.platform === "linux" && config.auth?.clientCertificate?.pinDialog?.enabled) {
      clientCertificate.initialize();
    }

    // Custom CA allowlist (issue #2762). Installed before the main window loads
    // so it covers the very first TLS handshake, and before profile partitions
    // exist so their sessions are caught by the session-created listener.
    certificateModule.installCertificateVerifyProc(config, app, session.defaultSession);

    await mainAppWindow.onAppReady(appConfig);

    // Flatpak only: record the background permission with the desktop portal
    // and set the Background Apps status line (issue #2815). No-op elsewhere.
    backgroundPortal.init();

    if (process.platform === "linux" && config.auth?.webauthn?.enabled) {
      await WebAuthn.initialize(mainAppWindow.getWindow(), config);
    }

    registerGlobalShortcuts(config, mainAppWindow, app);
    initializeAutoUpdater();

    // Attach the download manager after the partition is provisioned by the
    // main window. Issue #2512: without this, file downloads from Outlook are
    // silent and the user gets no completion feedback.
    downloadManager.initialize(session.fromPartition(config.partition));

    console.info('[IPC Security] Channel allowlisting enabled');
    console.info(`[IPC Security] ${allowedChannels.size} channels allowlisted`);
  } catch (error) {
    console.error('[STARTUP] Fatal error during app initialization:', { message: error.message, stack: error.stack });
    app.quit();
  }
}

function handleCertificateError(event, webContents, url, error, certificate, callback) {
  certificateModule.onAppCertificateError({
    event,
    webContents,
    url,
    error,
    certificate,
    callback,
    config,
  });
}

async function setBadgeCountHandler(_event, count) {
  if (!config.disableBadgeCount) {
    app.setBadgeCount(count);
    // Electron's own Linux badge path loads libunity at runtime, which none
    // of our packagings ship, so setBadgeCount is a silent no-op there.
    // Mirror the count over the LauncherEntry broadcast the docks actually
    // listen to — the same route downloads use for dock progress.
    if (os.platform() === "linux") {
      try {
        require("./downloadManager/launcherEntryEmitter").update({
          count,
          countVisible: count > 0,
        });
      } catch (error) {
        console.warn("[Badge] LauncherEntry emit failed", {
          message: error.message,
        });
      }
    }
  }
}

function handleGlobalShortcutDisabled() {
  for (const shortcut of config.disableGlobalShortcuts) {
    if (shortcut) {
      globalShortcut.register(shortcut, () => {
        console.debug(`Global shortcut ${shortcut} disabled`);
      });
    }
  }
}

function handleGlobalShortcutDisabledRevert() {
  for (const shortcut of config.disableGlobalShortcuts) {
    if (shortcut) {
      globalShortcut.unregister(shortcut);
    }
  }
}
