# Outlook for Linux

Unofficial Microsoft Outlook client for Linux. It wraps [Outlook on the web](https://outlook.office.com) (Microsoft 365 work/school accounts, or Outlook.com personal accounts) in a standalone Electron application.

> [!NOTE]
> Outlook for Linux is a community project. It is not affiliated with, endorsed by or supported by Microsoft. What works in Outlook on the web works here; what does not, does not.

## Features

- Outlook on the web in its own window, with the session kept between restarts
- Unread mail count on the tray icon and launcher badge (GNOME/Ubuntu Dock, KDE)
- Desktop notifications for new mail and calendar reminders
- `mailto:` link handler that opens a pre-filled Outlook compose window
- Outlook pop-outs (a message or event opened in a new window) stay in the app; other links open in your browser
- Microsoft sign-in support: Intune SSO (Linux identity broker), FIDO2 security keys, smartcard/PKCS#11 client certificates, custom CA certificates, NTLM/Basic auth, proxy support
- Download progress in KDE Plasma and on the dock icon
- Spell checking, zoom persistence, custom CSS, tray icon variants
- deb, rpm, tar.gz, AppImage (with in-app updates) and snap packages

## Installation

Download a package for your architecture (x64, arm64, armv7l) from the [Releases](https://github.com/Brumhilde/outlook-for-linux/releases) page:

```bash
# Debian / Ubuntu
sudo apt install ./outlook-for-linux_*_amd64.deb

# Fedora / RHEL / openSUSE
sudo dnf install ./outlook-for-linux-*.x86_64.rpm

# AppImage
chmod +x outlook-for-linux-*.AppImage && ./outlook-for-linux-*.AppImage
```

### Build from source

```bash
git clone https://github.com/Brumhilde/outlook-for-linux.git
cd outlook-for-linux
npm install
npm start                  # run from source
npm run dist:linux:x64     # build deb, rpm, tar.gz and AppImage into dist/
```

## Configuration

Settings are read from `~/.config/outlook-for-linux/config.json` (a system-wide `/etc/outlook-for-linux/config.json` is merged underneath it). The menu's **Settings → Open config file** creates and opens it. Most options apply after a restart. Every option is listed in [docs/configuration.md](docs/configuration.md).

### Outlook.com (personal accounts)

The app opens the Microsoft 365 work/school Outlook by default. For a personal Microsoft account:

```json
{
  "app": { "url": "https://outlook.live.com/mail/" }
}
```

### Unread count in other languages

The unread count is read from the Inbox entry in the folder pane. If Outlook is not in English, set the Inbox name it shows:

```json
{
  "unreadCount": { "folderLabel": "Posteingang" }
}
```

### mailto: links

Packages register Outlook for Linux as a `mailto:` handler. To make it the default:

```bash
xdg-mime default outlook-for-linux.desktop x-scheme-handler/mailto
```

A `mailto:` link opens a compose window with the recipients, subject and body filled in. If the app is still starting or waiting for you to sign in, the compose window opens once Outlook has loaded. Set `"mailto": { "enabled": false }` to turn this off.

### Useful options

| Option | Purpose |
| --- | --- |
| `window.closeOnCross` | Quit when the window is closed instead of hiding to the tray |
| `window.minimized` | Start minimized to the tray |
| `tray.iconType` | `default`, `light` or `dark` tray icon |
| `notificationMethod` | `web`, `electron` or `custom` notifications |
| `auth.intune.enabled` | Single sign-on through the Microsoft Identity Broker (Intune) |
| `auth.webauthn.enabled` | FIDO2 security keys on the Microsoft login page (needs `fido2-tools`) |
| `proxyServer` | `address:port` of an HTTP proxy |
| `appearance.cssLocation` | Path to a custom CSS file |

## Troubleshooting

- **Blank window:** try `"performance": { "disableGpu": true }` in the config file, or start with `--disableGpu`.
- **Signed out on every start:** make sure a Secret Service keyring (GNOME Keyring or KWallet) is running.
- **Logs:** start from a terminal with `ELECTRON_ENABLE_LOGGING=true outlook-for-linux --logConfig='{"transports":{"console":{"level":"debug"}}}'`.
- **Is it Outlook or the app?** Check whether the same thing happens in Outlook on the web in a browser.

## Credits

Outlook for Linux is a fork of [Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux) by Ismael Martinez and its contributors. The authentication, notification, tray, download and packaging code comes from that project.

## License

[GPL-3.0-or-later](LICENSE.md). Microsoft, Outlook and Microsoft 365 are trademarks of the Microsoft group of companies; this project uses no Microsoft logos.
