const { describe, it } = require("node:test");
const assert = require("node:assert");

const {
  isMailtoUrl,
  parseMailto,
  toComposeUrl,
  findMailtoArg,
} = require("../../app/mailto");

const M365 = "https://outlook.office.com/mail/";
const LIVE = "https://outlook.live.com/mail/";

describe("isMailtoUrl", () => {
  it("accepts mailto: URLs in any case", () => {
    assert.strictEqual(isMailtoUrl("mailto:a@example.com"), true);
    assert.strictEqual(isMailtoUrl("MAILTO:a@example.com"), true);
    assert.strictEqual(isMailtoUrl("mailto:"), true);
  });

  it("rejects everything else", () => {
    for (const value of ["https://example.com", "--mailto:a@b.c", "", null, undefined, 42]) {
      assert.strictEqual(isMailtoUrl(value), false, String(value));
    }
  });
});

describe("parseMailto", () => {
  it("reads recipients from the path and the to/cc/bcc fields", () => {
    const parsed = parseMailto(
      "mailto:a@example.com,b@example.com?to=c@example.com&cc=d@example.com;e@example.com&bcc=f@example.com",
    );
    assert.deepStrictEqual(parsed.to, ["a@example.com", "b@example.com", "c@example.com"]);
    assert.deepStrictEqual(parsed.cc, ["d@example.com", "e@example.com"]);
    assert.deepStrictEqual(parsed.bcc, ["f@example.com"]);
  });

  it("percent-decodes values and keeps a literal plus", () => {
    const parsed = parseMailto(
      "mailto:user+tag@example.com?subject=Hello%20there%20%26%20welcome&body=Line%201%0D%0ALine+2",
    );
    assert.deepStrictEqual(parsed.to, ["user+tag@example.com"]);
    assert.strictEqual(parsed.subject, "Hello there & welcome");
    assert.strictEqual(parsed.body, "Line 1\r\nLine+2");
  });

  it("matches field names case-insensitively and keeps the first subject", () => {
    const parsed = parseMailto("mailto:?Subject=First&SUBJECT=Second&CC=x@example.com");
    assert.strictEqual(parsed.subject, "First");
    assert.deepStrictEqual(parsed.cc, ["x@example.com"]);
    assert.deepStrictEqual(parsed.to, []);
  });

  it("tolerates malformed percent-encoding", () => {
    const parsed = parseMailto("mailto:a@example.com?subject=100%");
    assert.strictEqual(parsed.subject, "100%");
  });

  it("returns null for non-mailto input", () => {
    assert.strictEqual(parseMailto("https://example.com"), null);
  });
});

describe("toComposeUrl", () => {
  it("builds a Microsoft 365 compose deep link on the configured origin", () => {
    const url = toComposeUrl(
      "mailto:a@example.com,b@example.com?cc=c@example.com&subject=Hi%20there&body=See%20you",
      M365,
    );
    assert.strictEqual(
      url,
      "https://outlook.office.com/mail/deeplink/compose?to=a%40example.com%2Cb%40example.com&cc=c%40example.com&subject=Hi%20there&body=See%20you",
    );
  });

  it("uses the /mail/0/ compose route on outlook.live.com", () => {
    const url = toComposeUrl("mailto:a@example.com", LIVE);
    assert.strictEqual(url, "https://outlook.live.com/mail/0/deeplink/compose?to=a%40example.com");
  });

  it("follows other Outlook origins such as outlook.cloud.microsoft", () => {
    const url = toComposeUrl("mailto:a@example.com", "https://outlook.cloud.microsoft/mail/");
    assert.strictEqual(url, "https://outlook.cloud.microsoft/mail/deeplink/compose?to=a%40example.com");
  });

  it("encodes spaces as %20 rather than +", () => {
    const url = toComposeUrl("mailto:?subject=a%20b", M365);
    assert.ok(url.endsWith("?subject=a%20b"), url);
  });

  it("omits the query for an empty mailto:", () => {
    assert.strictEqual(toComposeUrl("mailto:", M365), "https://outlook.office.com/mail/deeplink/compose");
  });

  it("returns null for invalid input or a non-https base URL", () => {
    assert.strictEqual(toComposeUrl("https://example.com", M365), null);
    assert.strictEqual(toComposeUrl("mailto:a@example.com", "not a url"), null);
    assert.strictEqual(toComposeUrl("mailto:a@example.com", "http://outlook.office.com/"), null);
  });
});

describe("findMailtoArg", () => {
  it("finds the mailto: argument among other process arguments", () => {
    const args = ["/usr/bin/outlook-for-linux", "--no-sandbox", "mailto:a@example.com"];
    assert.strictEqual(findMailtoArg(args), "mailto:a@example.com");
  });

  it("returns null when there is none", () => {
    assert.strictEqual(findMailtoArg(["/usr/bin/outlook-for-linux"]), null);
    assert.strictEqual(findMailtoArg(undefined), null);
  });
});
