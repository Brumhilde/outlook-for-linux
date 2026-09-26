# Main App Window

Manages the primary BrowserWindow that hosts Outlook on the web.

## Components

- **[index.js](index.js)**: Entry point and window lifecycle management, auth recovery, link handling and mailto: compose windows
- **[browserWindowManager.js](browserWindowManager.js)**: Window creation and configuration

## Responsibilities

- Window state management (minimize, maximize, close)
- Web contents configuration and security settings
- Auth cookie maintenance and optional in-app re-authentication recovery (`auth.reauthRecovery.enabled`)
- Link handling: Outlook URLs opened with `window.open` (a message or event popped out) stay in-app as pop-out windows; all other links open in the system browser (or `urlHandling.defaultHandler`)

## mailto: Links

The app registers itself as the `mailto:` handler (`mailto.enabled`, on by default). A `mailto:` URL on the command line, either at launch or passed to the running instance through `second-instance`, is translated by [`app/mailto`](../mailto/index.js) into an Outlook compose deep link on the configured Outlook origin (`/mail/deeplink/compose`, or `/mail/0/deeplink/compose` on outlook.live.com). The link opens in its own compose window that shares the main window's session partition. Until the main window has loaded an Outlook page (for example, while the user is still signing in), the newest link waits and opens as soon as Outlook has loaded. Recipients, subject and body are never logged.
