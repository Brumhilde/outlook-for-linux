# Browser Module

Handles browser-side code injection and communication with the Outlook web page.

## Structure

- **[preload.js](preload.js)**: Preload script. Installs the `window.Notification` override, exposes `globalThis.electronAPI`, restores `File.path` on dropped/pasted files for Outlook hosts, forwards renderer errors, and loads the browser tools
- **[notifications/](notifications/)**: Bridges main-process notification click/close events back to the page
- **[tools/](tools/)**: Client-side scripts (unread count, tray badge, zoom, shortcuts, WebAuthn)

## Key Features

- Unread mail count tracking and tray icon updates
- Desktop notifications for new mail and calendar reminders (web, electron or custom toast)
- Keyboard shortcuts and zoom controls
- Pop-out windows (flagged with `POPOUT_WINDOW_ARG` from `app/helpers/popout.js`) skip the tray and unread-count tools
