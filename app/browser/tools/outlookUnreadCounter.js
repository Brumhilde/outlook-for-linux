// Tracks Outlook's unread mail count and dispatches it as an `unread-count`
// window event, which trayIconRenderer turns into the tray/dock badge.
//
// Outlook on the web does not reliably put the count in document.title, so
// two sources are tried in order:
//   1. document.title, matched against `unreadCount.titleRegex`
//   2. the folder pane: the tree item for `unreadCount.folderLabel` (the
//      Inbox) and the number shown next to it
// The Outlook DOM changes without notice, so every lookup is defensive and a
// miss simply keeps the last known count.

const DEFAULT_TITLE_REGEX = String.raw`\((\d+)\)`;
const DEFAULT_FOLDER_LABEL = "Inbox";
// DOM mutations in a mail client are constant; recount at most this often.
const RECOUNT_THROTTLE_MS = 1000;
// Safety net for changes the observer misses (e.g. a virtualised folder list
// swapping nodes in ways that do not touch the observed attributes).
const POLL_INTERVAL_MS = 30 * 1000;
const MAX_COUNT = 9999;

class OutlookUnreadCounter {
  #lastCount = -1;
  #titleRegex = null;
  #folderLabel = DEFAULT_FOLDER_LABEL;
  #recountTimer = null;

  init(config) {
    if (config?.unreadCount?.enabled === false) {
      console.debug("[UNREAD] Unread count disabled by config");
      return;
    }

    this.#titleRegex = this.#compileRegex(config?.unreadCount?.titleRegex);
    const label = config?.unreadCount?.folderLabel;
    this.#folderLabel =
      typeof label === "string" && label.trim() ? label.trim() : DEFAULT_FOLDER_LABEL;

    if (document.readyState === "loading") {
      globalThis.addEventListener("DOMContentLoaded", () => this.#start());
    } else {
      this.#start();
    }
  }

  #compileRegex(pattern) {
    const source = typeof pattern === "string" && pattern ? pattern : DEFAULT_TITLE_REGEX;
    try {
      return new RegExp(source);
    } catch {
      console.warn("[UNREAD] Invalid unreadCount.titleRegex, using the default");
      return new RegExp(DEFAULT_TITLE_REGEX);
    }
  }

  #start() {
    try {
      if (!globalThis.MutationObserver || !document.documentElement) {
        console.error("[UNREAD] Invalid DOM environment");
        return;
      }

      const observer = new globalThis.MutationObserver(() => this.#scheduleRecount());
      // The whole document, because both the <title> and the folder pane's
      // count badge can be replaced wholesale by Outlook's React renderer.
      observer.observe(document.documentElement, {
        childList: true,
        characterData: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["aria-label", "title"],
      });
      setInterval(() => this.#recount(), POLL_INTERVAL_MS);
      this.#recount();
      console.debug("[UNREAD] Unread counter attached");
    } catch (error) {
      console.error("[UNREAD] Failed to attach unread counter:", error.message);
    }
  }

  #scheduleRecount() {
    if (this.#recountTimer) return;
    this.#recountTimer = setTimeout(() => {
      this.#recountTimer = null;
      this.#recount();
    }, RECOUNT_THROTTLE_MS);
  }

  #recount() {
    try {
      const count = this.#countFromTitle() ?? this.#countFromFolderPane();
      if (count === null || count === this.#lastCount) return;
      this.#lastCount = count;
      globalThis.dispatchEvent(new CustomEvent("unread-count", { detail: { number: count } }));
    } catch (error) {
      console.debug("[UNREAD] Recount failed:", error.message);
    }
  }

  #countFromTitle() {
    const title = typeof document.title === "string" ? document.title.substring(0, 200) : "";
    const match = this.#titleRegex.exec(title);
    return match ? this.#toCount(match.slice(1).find(Boolean)) : null;
  }

  // Returns the Inbox unread count, 0 when the Inbox is visible with no
  // count, or null when the folder pane is not on screen (e.g. the calendar).
  #countFromFolderPane() {
    const item = this.#findFolderItem();
    if (!item) return null;

    // The count is usually its own element whose whole text is the number.
    for (const el of item.querySelectorAll("span, div")) {
      if (el.children.length === 0) {
        const text = el.textContent?.trim();
        if (text && /^\d+$/.test(text)) {
          return this.#toCount(text);
        }
      }
    }

    // Otherwise the accessible name carries it, e.g. "Inbox 3 unread".
    const labelText = (item.getAttribute("aria-label") || "").slice(this.#folderLabel.length);
    const digits = /\d+/.exec(labelText);
    return digits ? this.#toCount(digits[0]) : 0;
  }

  #findFolderItem() {
    const wanted = this.#folderLabel.toLowerCase();
    for (const item of document.querySelectorAll('[role="treeitem"]')) {
      const name = (item.getAttribute("title") || item.getAttribute("aria-label") || "")
        .trim()
        .toLowerCase();
      // The accessible name may append the count ("Inbox, 3 unread"), so a
      // prefix followed by a separator also matches; "Inboxes" does not.
      if (name === wanted || name.startsWith(wanted + " ") || name.startsWith(wanted + ",")) {
        return item;
      }
    }
    return null;
  }

  #toCount(value) {
    const number = Number.parseInt(value, 10);
    if (Number.isNaN(number) || number < 0) return null;
    return Math.min(number, MAX_COUNT);
  }
}

module.exports = new OutlookUnreadCounter();
