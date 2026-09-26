# Browser Tools

Client-side scripts loaded by the preload script into the Outlook on the web page to integrate it with the desktop.

## Overview

These tools run in the renderer process. They are loaded from `app/browser/preload.js` on `DOMContentLoaded` and initialised with the app configuration. Outlook pop-out windows (a message or event opened in its own window) load only the tools that make sense there; the tray and unread-count tools run in the main window only.

## Available Tools

#### [outlookUnreadCounter.js](outlookUnreadCounter.js)
Tracks the unread mail count and dispatches it as an `unread-count` window event. It first matches `document.title` against `unreadCount.titleRegex`; when the title carries no count it reads the number next to the `unreadCount.folderLabel` folder (the Inbox) in the folder pane. The Outlook DOM changes without notice, so every lookup is defensive and a miss keeps the last known count.

**Configuration**: `unreadCount.enabled`, `unreadCount.titleRegex`, `unreadCount.folderLabel` (set the latter to the Inbox name in your Outlook display language).

#### [trayIconChooser.js](trayIconChooser.js) & [trayIconRenderer.js](trayIconRenderer.js)
Pick the tray icon variant and draw the unread badge on it when an `unread-count` event arrives, then send `tray-update` and `set-badge-count` to the main process.

**Requires**: `ipcRenderer` passed during initialization (see CLAUDE.md, issue #1902).

#### [zoom.js](zoom.js)
Manages zoom level controls and persistence across sessions.

#### [shortcuts.js](shortcuts.js)
Implements in-app keyboard shortcuts for zoom and history navigation.

#### [emulatePlatform.js](emulatePlatform.js)
Reports a Windows platform to the page when `platform.emulateWindowsChromium` is set, for MFA flows that reject Linux.

#### [webauthnOverride.js](webauthnOverride.js)
Routes WebAuthn (FIDO2 security key) requests on the Microsoft login pages to the main process on Linux. See [../../webauthn/README.md](../../webauthn/README.md).

**Requires**: `ipcRenderer` passed during initialization.

## Global Shortcuts (Main Process)

System-wide keyboard shortcuts that work even when Outlook is not focused. When triggered, the keyboard event is forwarded to the Outlook window, which handles it with its built-in shortcuts. Configured via the `shortcuts.global` array in `config.json`.

**Disabled by default** - opt-in by adding shortcuts to your config:

```json
{
  "shortcuts": {
    "global": ["Control+Shift+M"]
  }
}
```

See [Electron Accelerators](https://www.electronjs.org/docs/latest/api/accelerator) for key combinations.

## Adding New Tools

1. Use camelCase file names and export an object with `init(config, ipcRenderer?)`
2. Add any configuration option to `app/config/options.js` and run `npm run generate-config-docs`
3. Register the tool in the `modules` array in `app/browser/preload.js` (and in `modulesRequiringIpc` if it needs `ipcRenderer`)
4. Use a `[TOOL_NAME]` log prefix and never log message content, addresses or subjects
5. Wrap DOM access in try/catch: the Outlook DOM can change unexpectedly
