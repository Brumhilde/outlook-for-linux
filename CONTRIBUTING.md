# Contributing to Outlook for Linux

Thank you for considering contributing!

## Quick Start

1. **Fork** the repository
2. **Clone** your fork and create a feature branch
3. **Make changes** (entry point: `app/index.js`)
4. **Test** your changes with `npm start`
5. **Submit** a pull request to the `main` branch

Each `app/` subfolder contains a README explaining its purpose.

## Development Setup

**Prerequisites:** Node.js 24 and npm.

```bash
git clone https://github.com/your-username/outlook-for-linux.git
cd outlook-for-linux
npm install

# Run from source
npm start

# Lint and unit tests (required before commits)
npm run lint
npm run test:unit

# End-to-end smoke tests (launch the app, check the Microsoft login redirect)
npm run test:e2e
```

After changing `app/config/options.js`, run `npm run generate-config-docs` and commit the regenerated `docs/configuration.md` and `docs/config-schema.json` (CI checks they are in sync). After adding or changing an IPC channel, add it to `app/security/ipcValidator.js` and run `npm run generate-ipc-docs`.

## Testing Pull Requests

Every pull request builds Linux packages (deb, rpm, tar.gz, AppImage) in GitHub Actions. Open the PR's **Checks** tab, pick the finished **Build & Release** run and download the artifact for your architecture from the **Artifacts** section.

## Upstream

Most of the platform code (authentication, Intune SSO, WebAuthn, notifications, downloads, packaging) comes from [Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux). Fixes to that shared code are often worth offering upstream too.
