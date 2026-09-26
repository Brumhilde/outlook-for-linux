const test = require("node:test");
const assert = require("node:assert");

const { isOutlookHost, isOutlookUrl, OUTLOOK_HOSTS } = require("../../app/helpers/outlookHosts");

test("isOutlookHost accepts the Outlook hosts, one subdomain label and the MCAS proxy", () => {
  for (const host of OUTLOOK_HOSTS) {
    assert.strictEqual(isOutlookHost(host), true, host);
    assert.strictEqual(isOutlookHost(`eu.${host}`), true, host);
    assert.strictEqual(isOutlookHost(`${host}.mcas.ms`), true, host);
  }
});

test("isOutlookHost declines look-alikes and non-strings", () => {
  for (const host of [
    "evil.com.outlook.office.com",
    "outlook.office.com.evil.com",
    "a.b.outlook.cloud.microsoft",
    "login.microsoftonline.com",
    "",
    undefined,
  ]) {
    assert.strictEqual(isOutlookHost(host), false, String(host));
  }
});

test("isOutlookUrl requires https on an Outlook host", () => {
  assert.strictEqual(isOutlookUrl("https://outlook.office.com/mail/"), true);
  assert.strictEqual(isOutlookUrl("https://outlook.live.com/mail/0/deeplink/compose"), true);
  assert.strictEqual(isOutlookUrl("http://outlook.office.com/mail/"), false);
  assert.strictEqual(isOutlookUrl("https://example.com/outlook.office.com"), false);
  assert.strictEqual(isOutlookUrl("about:blank"), false);
  assert.strictEqual(isOutlookUrl("not a url"), false);
});
