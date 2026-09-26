# Helpers

This directory contains various utility functions and helper modules used across different parts of the application. These functions are designed to perform common tasks, encapsulate reusable logic, and simplify code in other modules.
`index.js` exposes `HTTPHelper` (URL joining and `net`-based GET).

`outlookHosts.js` exports `isOutlookHost(hostname)`, `isOutlookUrl(url)` and the `OUTLOOK_HOSTS` list: whether a hostname is one of the hosts Outlook on the web is served from, an immediate subdomain of one, or its `.mcas.ms` (Defender for Cloud Apps) proxy. It has no Electron dependency, so both the main process (`mainAppWindow/index.js`) and the preload script use it.

`popout.js` exports `POPOUT_WINDOW_ARG`, the switch that marks in-app Outlook pop-out windows so the preload skips the tray and unread-count tools there.
