# Security Policy

## Supported Versions

Only the latest release receives security updates. Upgrade to the latest version to get fixes.

## Reporting a Vulnerability

If you discover a security vulnerability in Outlook for Linux, please report it responsibly:

1. **Do not** open a public GitHub issue for security vulnerabilities.
2. Use GitHub's private vulnerability reporting (the repository's **Security → Report a vulnerability** tab) to submit details confidentially.
3. Include a description of the vulnerability, steps to reproduce and any potential impact.

Vulnerabilities in code inherited unchanged from [Teams for Linux](https://github.com/IsmaelMartinez/teams-for-linux) should also be reported upstream.

## Privacy & Data Protection

See the [Privacy & Data Protection statement](PRIVACY.md).

## Security Architecture

- The main window runs with `contextIsolation: false` (the preload installs the notification bridge directly into the page); every IPC channel is checked against the allowlist in `app/security/ipcValidator.js`, and payloads are sanitised against prototype pollution.
- mailto: compose windows run with `contextIsolation: true` and `sandbox: true` and no preload.
- Logs are passed through a PII sanitizer (`app/utils/logSanitizer.js`); message content, addresses and subjects are never logged.
