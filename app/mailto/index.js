// Translates mailto: links (RFC 6068) into Outlook on the web compose deep
// links. Pure functions with no Electron imports so they can be unit tested.

const PROTOCOL = "mailto";

// Header fields Outlook's compose deep link understands. Anything else in a
// mailto: query (In-Reply-To, custom headers) is dropped.
const COMPOSE_FIELDS = ["to", "cc", "bcc", "subject", "body"];

/**
 * @param {unknown} value - A command line argument or URL
 * @returns {boolean} Whether it is a mailto: URL
 */
function isMailtoUrl(value) {
  return typeof value === "string" && /^mailto:/i.test(value.trim());
}

// RFC 6068 uses plain percent-encoding: unlike form encoding, "+" is a
// literal plus (common in addresses such as user+tag@example.com).
function decode(value) {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function splitAddresses(value) {
  return value
    .split(/[,;]/)
    .map((address) => address.trim())
    .filter(Boolean);
}

/**
 * Parses a mailto: URL into its recipients and compose fields.
 *
 * @param {string} mailtoUrl - e.g. "mailto:a@example.com?cc=b@example.com&subject=Hi"
 * @returns {{to: string[], cc: string[], bcc: string[], subject: string, body: string}|null}
 *   null when the argument is not a mailto: URL
 */
function parseMailto(mailtoUrl) {
  if (!isMailtoUrl(mailtoUrl)) {
    return null;
  }

  const raw = mailtoUrl.trim().slice(PROTOCOL.length + 1);
  const queryStart = raw.indexOf("?");
  const path = queryStart === -1 ? raw : raw.slice(0, queryStart);
  const query = queryStart === -1 ? "" : raw.slice(queryStart + 1);

  const result = { to: [], cc: [], bcc: [], subject: "", body: "" };
  result.to.push(...splitAddresses(decode(path)));

  for (const pair of query.split("&")) {
    if (!pair) continue;
    const eq = pair.indexOf("=");
    const name = decode(eq === -1 ? pair : pair.slice(0, eq)).toLowerCase();
    const value = eq === -1 ? "" : decode(pair.slice(eq + 1));

    if (name === "to" || name === "cc" || name === "bcc") {
      result[name].push(...splitAddresses(value));
    } else if (name === "subject" || name === "body") {
      // RFC 6068 leaves repeated fields undefined; keep the first.
      result[name] ||= value;
    }
  }

  return result;
}

/**
 * Builds the Outlook compose deep link for a mailto: URL.
 *
 * The origin comes from the configured app URL, so the link opens in the same
 * Outlook the user is signed in to. Outlook.com (consumer) keeps its compose
 * route under the mailbox index, `/mail/0/`.
 *
 * @param {string} mailtoUrl - The mailto: URL to translate
 * @param {string} baseUrl - The configured Outlook URL (config.url)
 * @returns {string|null} The compose URL, or null if either input is invalid
 */
function toComposeUrl(mailtoUrl, baseUrl) {
  const parsed = parseMailto(mailtoUrl);
  if (!parsed) {
    return null;
  }

  let base;
  try {
    base = new URL(baseUrl);
  } catch {
    return null;
  }
  if (base.protocol !== "https:") {
    return null;
  }

  const composePath = base.hostname.endsWith("outlook.live.com")
    ? "/mail/0/deeplink/compose"
    : "/mail/deeplink/compose";
  // Encoded by hand rather than with URLSearchParams, which writes spaces as
  // "+"; percent-encoding keeps a subject's spaces from showing up as pluses.
  const query = COMPOSE_FIELDS.map((field) => {
    const value = Array.isArray(parsed[field]) ? parsed[field].join(",") : parsed[field];
    return value ? `${field}=${encodeURIComponent(value)}` : null;
  }).filter(Boolean);

  const composeUrl = new URL(composePath, base.origin).toString();
  return query.length ? `${composeUrl}?${query.join("&")}` : composeUrl;
}

/**
 * Finds the first mailto: URL in a list of command line arguments.
 *
 * @param {string[]} args - process.argv or second-instance argv
 * @returns {string|null}
 */
function findMailtoArg(args) {
  return (args || []).find(isMailtoUrl) ?? null;
}

module.exports = {
  PROTOCOL,
  isMailtoUrl,
  parseMailto,
  toComposeUrl,
  findMailtoArg,
};
