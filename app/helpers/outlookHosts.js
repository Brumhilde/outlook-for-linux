// The hosts Outlook on the web is served from. Tenants behind Defender for
// Cloud Apps load it at `<host>.mcas.ms`, which counts as the underlying host.
const OUTLOOK_HOSTS = [
  "outlook.office.com",
  "outlook.office365.com",
  "outlook.cloud.microsoft",
  "outlook.live.com",
];
const MCAS_SUFFIX = ".mcas.ms";

/**
 * Whether a hostname is one of the Outlook hosts or an immediate subdomain of
 * one. Only one label is allowed in front, so `evil.com.outlook.office.com`
 * does not pass.
 *
 * @param {string} hostname - Lower-case hostname, as `URL.hostname` yields it
 * @returns {boolean}
 */
function isOutlookHost(hostname) {
  if (typeof hostname !== "string") {
    return false;
  }
  if (hostname.endsWith(MCAS_SUFFIX)) {
    hostname = hostname.slice(0, -MCAS_SUFFIX.length);
  }
  return OUTLOOK_HOSTS.some(
    (domain) =>
      hostname === domain ||
      (hostname.endsWith("." + domain) &&
        !hostname.slice(0, -(domain.length + 1)).includes(".")),
  );
}

/**
 * Whether a URL is an https URL on one of the Outlook hosts.
 *
 * @param {string} url
 * @returns {boolean}
 */
function isOutlookUrl(url) {
  try {
    const parsed = new URL(url);
    return parsed.protocol === "https:" && isOutlookHost(parsed.hostname);
  } catch {
    return false;
  }
}

module.exports = { OUTLOOK_HOSTS, isOutlookHost, isOutlookUrl };
