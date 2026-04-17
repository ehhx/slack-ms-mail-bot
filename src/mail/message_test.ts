import { assertEquals } from "@std/assert";
import {
  attachmentSummaryText,
  buildDedupeKey,
  detectVerificationCode,
  htmlToPlainText,
  notificationBodyText,
  toPreviewText,
} from "./message.ts";

Deno.test("buildDedupeKey prefers internetMessageId", () => {
  assertEquals(
    buildDedupeKey("mailbox-1", {
      messageId: "graph-id",
      internetMessageId: "internet-id",
      subject: "Hello",
    }),
    "internet-id",
  );
});

Deno.test("buildDedupeKey falls back to mailbox/message id", () => {
  assertEquals(
    buildDedupeKey("mailbox-1", {
      messageId: "graph-id",
      subject: "Hello",
    }),
    "mailbox-1:graph-id",
  );
});

Deno.test("toPreviewText normalizes whitespace and truncates", () => {
  assertEquals(toPreviewText("hello\n\nworld", 10), "hello\n\nwo…");
});

Deno.test("htmlToPlainText strips markup and decodes entities", () => {
  assertEquals(
    htmlToPlainText(
      "<p>Hello&nbsp;<strong>world</strong><br>Tom &amp; Jerry</p>",
    ),
    "Hello world\nTom & Jerry",
  );
});

Deno.test("htmlToPlainText preserves anchor links and image markers", () => {
  assertEquals(
    htmlToPlainText(
      '<p><a href="https://example.com">Read more</a><br><img src="cid:image001" alt="banner"></p>',
    ),
    "Read more (https://example.com)\n[内联图片：banner]",
  );
});

Deno.test("notificationBodyText prefers bodyText and converts html to text", () => {
  assertEquals(
    notificationBodyText({
      messageId: "msg-1",
      subject: "Hello",
      bodyText: "<div>Line 1<br>Line 2</div>",
      bodyContentType: "html",
    }),
    "Line 1\nLine 2",
  );
});

Deno.test("notificationBodyText falls back to image-only hint", () => {
  assertEquals(
    notificationBodyText({
      messageId: "msg-1",
      subject: "Hello",
      bodyText: '<img src="cid:image001">',
      bodyContentType: "html",
      attachments: [
        { name: "image001.png", contentType: "image/png", isInline: true },
      ],
    }),
    "[内联图片]",
  );
});

Deno.test("detectVerificationCode prefers body code over subject brand word", () => {
  assertEquals(
    detectVerificationCode({
      subject: "Your login code for Nous Research",
      body:
        "Log in to Nous Research\nYour code is\n123522\nThis code expires in 10 minutes.",
    }),
    "123522",
  );
});

Deno.test("detectVerificationCode still detects subject short code when subject contains digits", () => {
  assertEquals(
    detectVerificationCode({
      subject: "Your ChatGPT code is 253225",
    }),
    "253225",
  );
});

Deno.test("attachmentSummaryText formats attachment list", () => {
  assertEquals(
    attachmentSummaryText([
      {
        name: "diagram.png",
        contentType: "image/png",
        size: 2048,
        isInline: true,
      },
      { name: "spec.pdf", contentType: "application/pdf", size: 10 * 1024 },
    ]),
    "*附件* (2)\n• diagram.png (inline, image/png, 2 KB)\n• spec.pdf (application/pdf, 10 KB)",
  );
});
