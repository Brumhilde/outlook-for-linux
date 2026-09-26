const { ipcRenderer } = require("electron");
const { POPOUT_WINDOW_ARG } = require("../helpers/popout");

// Outlook pop-out windows (a message or event opened in its own window) share
// this preload. Only the main window may drive the tray icon and badge.
const isPopoutWindow = process.argv.includes(POPOUT_WINDOW_ARG);

// #2677: Electron removed the non-standard `File.path` from dropped files.
// Restore it via webUtils.getPathForFile before the page's drop handler reads
// it, scoped to Outlook hosts so the SSO/auth pages this window also loads
// can't read local paths off dropped files.
//
// The same stripping hits pasted files: when a user copies a file in their
// file manager and pastes it into a message, Chromium surfaces it as a File on
// the paste event's clipboardData, so restore the path there too. Restore the path on a
// capture-phase paste listener too. Raw image-bit paste (screenshots) arrives
// as a Blob with no path and is unaffected.
try {
  const { webUtils } = require("electron");
  const { isOutlookHost } = require("../helpers/outlookHosts");
  // Restore the non-standard `File.path` on every File in a FileList, in place.
  // No-op for blob-backed files (screenshots) since webUtils only resolves a
  // path for files that originated from the OS file list; those are left as-is.
  const restoreFilePaths = (files) => {
    if (!files?.length) {
      return;
    }
    for (const file of files) {
      if (file.path) {
        continue;
      }
      try {
        const path = webUtils.getPathForFile(file);
        if (path) {
          Object.defineProperty(file, "path", {
            value: path,
            writable: true,
            enumerable: true,
            configurable: true,
          });
        }
      } catch {
        // leave the file untouched if the path can't be resolved
      }
    }
  };
  globalThis.addEventListener(
    "drop",
    (event) => {
      if (!isOutlookHost(globalThis.location.hostname)) {
        return;
      }
      restoreFilePaths(event.dataTransfer?.files);
    },
    true,
  );
  globalThis.addEventListener(
    "paste",
    (event) => {
      if (!isOutlookHost(globalThis.location.hostname)) {
        return;
      }
      restoreFilePaths(event.clipboardData?.files);
    },
    true,
  );
} catch {
  // webUtils unavailable
}

// Note: IPC validation handled by main process, no need for duplicate validation here
globalThis.electronAPI = {
  send: (channel, ...args) => {
    return ipcRenderer.send(channel, ...args);
  },

  getConfig: () => ipcRenderer.invoke("get-config"),

  showNotification: (options) => {
    if (!options || typeof options !== 'object') {
      return Promise.reject(new Error('Invalid notification options'));
    }
    return ipcRenderer.invoke("show-notification", options);
  },
  playNotificationSound: (options) => {
    if (options && typeof options !== 'object') {
      return Promise.reject(new Error('Invalid sound options'));
    }
    return ipcRenderer.invoke("play-notification-sound", options);
  },
  sendNotificationToast: (data) => {
    if (!data || typeof data !== 'object') {
      throw new Error('Invalid notification toast data');
    }
    ipcRenderer.send("notification-show-toast", data);
  },

  setBadgeCount: (count) => {
    if (typeof count !== 'number' || count < 0 || count > 9999) {
      console.error('Invalid badge count:', count);
      return Promise.reject(new Error('Invalid badge count'));
    }
    return ipcRenderer.invoke("set-badge-count", count);
  },

  updateTray: (icon, flash) => {
    return ipcRenderer.send("tray-update", { icon, flash });
  },

  getZoomLevel: (partition) => {
    if (typeof partition !== 'string' || partition.length > 100) {
      return Promise.reject(new Error('Invalid partition'));
    }
    return ipcRenderer.invoke("get-zoom-level", partition);
  },
  saveZoomLevel: (data) => {
    if (!data || typeof data !== 'object' || typeof data.level !== 'number') {
      return Promise.reject(new Error('Invalid zoom data'));
    }
    return ipcRenderer.invoke("save-zoom-level", data);
  },

  navigateBack: () => ipcRenderer.send("navigate-back"),
  navigateForward: () => ipcRenderer.send("navigate-forward"),
  getNavigationState: () => ipcRenderer.invoke("get-navigation-state"),
  onNavigationStateChanged: (callback) => {
    if (typeof callback !== 'function') {
      console.error('Invalid callback for navigation state changed');
      return;
    }
    return ipcRenderer.on("navigation-state-changed", callback);
  },

  sessionType: process.env.XDG_SESSION_TYPE || "x11",
};

// Config is fetched asynchronously; the Notification override below reads it via closure
let notificationConfig = null;
ipcRenderer.invoke("get-config").then((config) => {
  notificationConfig = config;
  console.debug("Preload: Config loaded for notifications:", {
    notificationMethod: config?.notificationMethod,
    disableNotifications: config?.disableNotifications
  });
}).catch((err) => {
  console.error("Preload: Failed to load config for notifications:", err);
});

// Create a Notification-like stub so the page can manage lifecycle without errors.
// Without addEventListener/close/dispatchEvent, a web app's notification state
// machine can break after the first notification, stopping later ones.
function createNotificationStub() {
  const stub = {
    onclick: null,
    onclose: null,
    onerror: null,
    onshow: null,
    close() { if (this.onclose) this.onclose(); },
    addEventListener(type, listener) {
      if (type === 'click') this.onclick = listener;
      else if (type === 'close') this.onclose = listener;
      else if (type === 'show') this.onshow = listener;
      else if (type === 'error') this.onerror = listener;
    },
    removeEventListener(type, listener) {
      if (type === 'click' && (!listener || this.onclick === listener)) this.onclick = null;
      else if (type === 'close' && (!listener || this.onclose === listener)) this.onclose = null;
      else if (type === 'show' && (!listener || this.onshow === listener)) this.onshow = null;
      else if (type === 'error' && (!listener || this.onerror === listener)) this.onerror = null;
    },
    dispatchEvent() { return true; },
  };
  // Fire the show event asynchronously like a real Notification
  setTimeout(() => { if (stub.onshow) stub.onshow(); }, 0);
  return stub;
}

function playNotificationSound(notifSound) {
  // Skip renderer-side sound for "electron" method — the main process
  // notification service already plays the sound before showing the notification.
  const method = notificationConfig?.notificationMethod || "web";
  if (method === "electron") {
    return;
  }
  if (globalThis.electronAPI?.playNotificationSound) {
    try {
      console.debug("Requesting application to play sound");
      globalThis.electronAPI.playNotificationSound(notifSound);
    } catch (e) {
      console.debug("playNotificationSound failed", e);
    }
  }
}

function createWebNotification(classicNotification, title, options) {
  const notifSound = {
    type: options.type,
    audio: "default",
    title: title,
    body: options.body,
  };
  playNotificationSound(notifSound);

  // Return actual native notification object (the page manages its lifecycle)
  console.debug("Continues to default notification workflow");
  if (classicNotification) {
    try {
      return new classicNotification(title, options);
    } catch (err) {
      console.debug("Could not create native notification:", err);
      return null;
    }
  }
  return null;
}

// Bridges the main process's notification lifecycle onto the stubs above: close,
// so Outlook knows when the system dismissed a notification, and click, so
// Outlook's own handler can open the message or event it came from (issue #2768).
// Two ipcRenderer listeners for the whole renderer, not a pair per notification.
const NotificationBridge = require("./notifications/notificationBridge");
const notificationBridge = new NotificationBridge(ipcRenderer);

function createElectronNotification(options) {
  const notificationId = crypto.randomUUID();
  const stub = createNotificationStub();
  if (globalThis.electronAPI?.showNotification) {
    // Register before invoking: main can emit click or close as soon as it has
    // shown the notification. stub.close() from the page and the notification-closed
    // IPC both route through the bridge, so whichever lands first fires onclose
    // exactly once. Without electronAPI the stub keeps its own close().
    notificationBridge.register(notificationId, stub);
    stub.close = () => notificationBridge.close(notificationId);
    globalThis.electronAPI
      .showNotification({ ...options, notificationId })
      .catch((e) => {
        console.debug("showNotification failed", e);
      });
  }
  return stub;
}

function createCustomNotification(title, options) {
  const notificationData = {
    id: crypto.randomUUID(),
    timestamp: Date.now(),
    title: title,
    body: options.body || '',
    icon: options.icon,
  };

  const notifSound = {
    type: options.type,
    audio: "default",
    title: title,
    body: options.body,
  };
  playNotificationSound(notifSound);

  try {
    if (globalThis.electronAPI?.sendNotificationToast) {
      globalThis.electronAPI.sendNotificationToast(notificationData);
    } else {
      console.warn("sendNotificationToast API not available");
    }
  } catch (e) {
    console.error("Failed to send custom notification:", e);
  }

  return createNotificationStub();
}

// Override window.Notification immediately before Outlook loads
// Using factory function pattern instead of class to avoid "return in constructor" anti-pattern
(function() {
  const ICON_BASE64 =
    "data:image/png;base64,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";

  const classicNotification = globalThis.Notification;

  // Factory function that creates notification objects (avoids "return in constructor" issue)
  function CustomNotification(title, options) {
    // Use config from closure scope (will be null initially, populated async)
    if (notificationConfig?.disableNotifications) {
      // Return dummy object to avoid page errors
      return { onclick: null, onclose: null, onerror: null };
    }

    options = options || {};
    options.icon = options.icon || ICON_BASE64;
    options.title = options.title || title;
    options.type = options.type || "new-message";
    // Default Ubuntu Unity DE auto-closes. Users on GNOME and similar can opt
    // into persistent notifications via `notifications.timeoutType: "never"`
    // (issue #2411). Mirrors Electron's Notification timeoutType.
    options.timeoutType =
      notificationConfig?.notifications?.timeoutType === "never"
        ? "never"
        : "default";
    options.requireInteraction = options.timeoutType === "never";

    // Default to "web" if config not loaded yet
    const method = notificationConfig?.notificationMethod || "web";

    if (method === "custom") {
      return createCustomNotification(title, options);
    }

    if (method === "web") {
      const notification = createWebNotification(classicNotification, title, options);
      return notification || { onclick: null, onclose: null, onerror: null };
    }

    return createElectronNotification(options);
  }

  CustomNotification.requestPermission = async function() {
    return "granted";
  };

  Object.defineProperty(CustomNotification, 'permission', {
    get: function() {
      return "granted";
    }
  });

  globalThis.Notification = CustomNotification;
  console.debug("Preload: CustomNotification factory initialized");
})();

document.addEventListener('DOMContentLoaded', async () => {
  console.debug("Preload: DOMContentLoaded, initializing browser modules...");
  try {
    const config = await ipcRenderer.invoke("get-config");
    console.debug("Preload: Got config:", {
      trayIconEnabled: config?.trayIconEnabled,
      isPopoutWindow,
    });
    
    // Only the main window drives the tray icon and badge; pop-outs
    // would otherwise report their own (empty) unread state.
    const modules = [
      { name: "zoom", path: "./tools/zoom" },
      { name: "shortcuts", path: "./tools/shortcuts" },
      { name: "emulatePlatform", path: "./tools/emulatePlatform" },
      { name: "webauthnOverride", path: "./tools/webauthnOverride" },
      ...(isPopoutWindow
        ? []
        : [
            { name: "trayIconRenderer", path: "./tools/trayIconRenderer" },
            { name: "outlookUnreadCounter", path: "./tools/outlookUnreadCounter" },
          ]),
    ];

    // CRITICAL: These modules need ipcRenderer for IPC communication (see CLAUDE.md)
    const modulesRequiringIpc = new Set(["trayIconRenderer", "webauthnOverride"]);

    let successCount = 0;
    for (const module of modules) {
      try {
        const moduleInstance = require(module.path);
        if (modulesRequiringIpc.has(module.name)) {
          moduleInstance.init(config, ipcRenderer);
        } else {
          moduleInstance.init(config);
        }
        successCount++;
      } catch (err) {
        console.error(`Preload: Failed to load ${module.name}:`, err.message);
      }
    }
    
    console.info(`Preload: ${successCount}/${modules.length} browser modules initialized successfully`);

    // Listen for config changes from the main process (e.g., when menu toggles are clicked)
    ipcRenderer.on("config-changed", (_event, configChanges) => {
      for (const [key, value] of Object.entries(configChanges)) {
        config[key] = value;
      }
    });

  } catch (error) {
    console.error("Preload: Failed to initialize browser modules:", error);
  }
});

// Forward unhandled promise rejections and window errors to main for diagnostics.
// Plain objects without a `.message` (and `undefined` rejections) previously stringified to
// the literals "[object Object]" / "undefined", which discarded all diagnostic content.
function serializeRejectionReason(reason) {
  // The whole body is wrapped in try/catch so a throwing `reason.message`
  // getter (or any other unexpected exception) degrades to a sentinel
  // string instead of propagating to the outer handler and dropping the
  // whole rejection payload.
  try {
    if (reason === undefined) return "<undefined>";
    if (reason === null) return "<null>";
    if (typeof reason === "string") return reason;
    if (typeof reason !== "object") return String(reason);
    if (typeof reason.message === "string" && reason.message.length > 0) return reason.message;
    const seen = new WeakSet();
    return JSON.stringify(reason, (_key, value) => {
      if (typeof value === "object" && value !== null) {
        if (seen.has(value)) return "[Circular]";
        seen.add(value);
      }
      return value;
    }) ?? "[unserializable rejection]";
  } catch {
    return "[unserializable rejection]";
  }
}

try {
  globalThis.addEventListener("unhandledrejection", (event) => {
    try {
      const reason = event?.reason;
      const errorData = {
        message: serializeRejectionReason(reason).substring(0, 1000),
        stack: reason?.stack ? String(reason.stack).substring(0, 5000) : null,
        timestamp: Date.now(),
      };

      ipcRenderer.send("unhandled-rejection", errorData);
    } catch (err) {
      console.debug("Unhandled rejection forwarding failed:", err);
      // Best-effort forwarding, never throw from preload
    }
  });

  globalThis.addEventListener("error", (event) => {
    try {
      const errorData = {
        message: event?.message ? String(event.message).substring(0, 1000) : '',
        filename: event?.filename ? String(event.filename).substring(0, 200) : '',
        lineno: typeof event?.lineno === 'number' ? event.lineno : 0,
        colno: typeof event?.colno === 'number' ? event.colno : 0,
        timestamp: Date.now(),
        errorStack: event?.error?.stack ? String(event.error.stack).substring(0, 5000) : null,
      };
      
      ipcRenderer.send("window-error", errorData);
    } catch (err) {
      console.debug("Window error forwarding failed:", err);
    }
  });
} catch (err) {
  console.debug("Error handler setup failed:", err);
}
