// Command line switch the main process adds to Outlook pop-out windows (a
// message or event opened in its own window). Pop-outs inherit the main
// window's preload; the preload checks for this switch so that only the main
// window drives the tray icon and unread count.
const POPOUT_WINDOW_ARG = "--outlook-for-linux-popout";

module.exports = { POPOUT_WINDOW_ARG };
