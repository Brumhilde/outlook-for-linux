# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Essential Commands

**Development:**
- `npm start` - Run application in development mode with trace warnings
- `npm run lint` - Run ESLint validation (mandatory before commits)
- `npm run test:unit` - Fast unit suite (`node --test 'tests/unit/*.test.js'`); run before every commit
- `npm run test:e2e` - Run end-to-end tests with Playwright

**Building:**
- `npm run pack` - Development build without packaging
- `npm run dist:linux` - Build Linux packages (AppImage, deb, rpm, tar.gz)
- `npm run dist:linux:x64` - Build the x64 Linux packages CI publishes

**Utility:**
- `npm run generate-release-info` - Generate release information file (reads `outlook-for-linux.appdata.xml`)
- `npm run generate-ipc-docs` - Regenerate `docs/ipc-api.md` from IPC registration comments
- `npm run generate-config-docs` - Regenerate `docs/configuration.md` and `docs/config-schema.json`; mandatory after editing `app/config/options.js` (CI has a drift guard)

**Release:** managed by release-please. Merging the Release PR produces a **draft** GitHub release that is published by hand.

## Project Architecture

Outlook for Linux is an Electron application that wraps Outlook on the web (`https://outlook.office.com/mail/` by default, configurable via `app.url`). It is a fork of [Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux) with the Teams-only features (calls, screen sharing, MQTT, custom backgrounds, stickers, quick chat, multi-account profiles) removed.

**Key file locations:**
- **Entry Point:** `app/index.js` - Main Electron process
- **Startup:** `app/startup/` - Command line switches
- **Configuration:** `app/config/options.js` (schema, single source of truth) and `app/appConfiguration/`
- **Main Window:** `app/mainAppWindow/` - BrowserWindow, auth recovery, link handling, mailto: compose windows
- **mailto:** `app/mailto/` - Pure mailto: → Outlook compose deep-link translation
- **Browser Tools:** `app/browser/tools/` - Scripts loaded by the preload into the Outlook page
- **Host lists:** `app/helpers/outlookHosts.js` - Hosts treated as Outlook (in-app pop-outs, drag-drop paths, CSP exemption)

Module-specific README.md files in `app/` subdirectories describe each module.

## Development Patterns

### Code Style Requirements
- **NO `var`** - Use `const` by default, `let` for reassignment
- **Strict equality** - `===` / `!==` (ESLint `eqeqeq`)
- **async/await** - Use instead of promise chains
- **Private fields** - Use JavaScript `#property` syntax for class private members
- **Arrow functions** - For concise callbacks

### Configuration Management
- All options are declared in `app/config/options.js`; run `npm run generate-config-docs` after changes
- Flat option names are deprecated aliases of nested ones (`app/config/renames.js`); modules read the flat name
- yargs replaces object options wholesale, so code must tolerate missing leaves (`config.mailto?.enabled !== false`)

### IPC Communication
- Use `ipcMain.handle` for request-response patterns and `ipcMain.on` for fire-and-forget
- Add a descriptive comment above each IPC channel registration
- Add every channel to the allowlist in `app/security/ipcValidator.js`
- Run `npm run generate-ipc-docs` after adding/modifying IPC channels

### Logging Guidelines

**CRITICAL: PII Protection.** Never log personally identifiable information:
- Email addresses, recipients, subjects, message bodies, mailto: URLs
- Usernames, account IDs, tokens, credentials
- URL query parameters (may contain tokens or recipients)
- Certificate details, SSO/Intune account information

Use structured, PII-free logs: `console.info('[MAILTO] mailto: link received')`. Levels: `error` for errors needing attention, `warn` for potential issues, `info` for key state changes, `debug` sparingly.

## Critical Module Initialization Requirements

**DO NOT REMOVE** `trayIconRenderer` (and `webauthnOverride`) from `modulesRequiringIpc` in `app/browser/preload.js`. Without `ipcRenderer` the tray badge silently stops updating (upstream issue #1902). `tests/unit/preloadModules.test.js` guards this.

Pop-out windows are flagged with `POPOUT_WINDOW_ARG` (`app/helpers/popout.js`); the preload skips the tray and unread-count tools there so only the main window drives the badge.

## Testing and Quality

- **Unit tests** (`tests/unit/`, `node:test`, no DOM): tests for injected browser scripts assert on source text. To verify renderer behaviour, run a throwaway `node_modules/.bin/electron probe.js` with a hidden `BrowserWindow` and `executeJavaScript`.
- **E2E tests** (`tests/e2e/`, Playwright): each test uses a temporary userData dir via `E2E_USER_DATA_DIR` and checks the launch → Microsoft login redirect. `tests/e2e/authenticated/` needs a signed-in session dir in `E2E_SESSION_DIR`.
- Run `npm run lint` and `npm run test:unit` before every commit.

## Important Notes

- Browser scripts must be defensive: the Outlook DOM changes without notice. The unread counter (`app/browser/tools/outlookUnreadCounter.js`) reads the page title first, then the Inbox entry in the folder pane.
- New functionality belongs in its own module, not in `app/index.js`.
- Update module README.md files alongside code changes.
- Linux is the primary platform.
