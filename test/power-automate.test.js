// Exercises the paste-ready Office Script entirely with synthetic Outlook content.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import { runInNewContext } from "node:vm";

const script = readFileSync(
  new URL("../power-automate/format-staff-news.ts", import.meta.url),
  "utf8",
);
const flow = JSON.parse(
  readFileSync(new URL("../power-automate/flow-definition.json", import.meta.url), "utf8"),
);
const format = runInNewContext(stripTypeScriptTypes(script) + "\nmain;");
const workbook = new Proxy(
  {},
  { get: () => assert.fail("The script should not access the workbook") },
);
const mailbox = "publisher@bond.edu.au";
const logo = readFileSync(new URL("../assets/bond-logo.png", import.meta.url));
const careers = readFileSync(new URL("../assets/careers-at-bond.png", import.meta.url));
// Output captured from the existing Staff News layout for this synthetic message.
const expectedHtml = readFileSync(new URL("./fixtures/staff-news.html", import.meta.url), "utf8");
const bondStory = "https://bonduni.sharepoint.com/sites/news/SitePages/Example.aspx";
const careersPage =
  "https://bonduni.sharepoint.com/sites/ShowcaseNews/SitePages/Careers-at-Bond(3).aspx";
const internalFooter = "This email is generated through the organisation's use of Microsoft 365.";
const image = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aEHoAAAAASUVORK5CYII=",
  "base64",
);
const outlookAttachments = [
  {
    "@odata.type": "#microsoft.graph.fileAttachment",
    isInline: true,
    contentId: "<Photo@Outlook>",
    contentType: "image/png",
    contentBytes: image.toString("base64"),
  },
];

function sourceMail({ href = bondStory, notice = internalFooter, extra = "" } = {}) {
  return `<html lang="en"><body>
    <table><tr><td><img src="cid:Photo%40Outlook" alt="People &amp; Places" width="200"></td>
    <td><div><a href="${href}">People &amp; Places</a></div>
    <div>Updates to share. &#x1F4E2;</div><div>By Staff Writer</div>
    <div>25/11/2026</div><div>42 views</div></td></tr>
    ${extra}</table><p>${notice}</p></body></html>`;
}

function draft(html = sourceMail(), date = "2026-11-26", careersWeekday) {
  return JSON.parse(
    format(workbook, html, JSON.stringify(outlookAttachments), mailbox, date, careersWeekday),
  );
}

test("the Office Script generates the existing Staff News layout and binds the inline images to a draft", () => {
  const result = draft();
  assert.equal(result.subject, "Bond Staff News | Thursday 26 November 2026");
  assert.equal(result.body.contentType, "HTML");
  assert.equal(result.body.content, expectedHtml);
  assert.deepEqual(result.toRecipients, [{ emailAddress: { address: mailbox } }]);
  assert.equal(result.attachments.length, 2);
  for (const attachment of result.attachments) {
    assert.equal(attachment["@odata.type"], "#microsoft.graph.fileAttachment");
    assert.equal(attachment.isInline, true);
    assert.ok(result.body.content.includes("cid:" + attachment.contentId));
  }
  assert.deepEqual(Buffer.from(result.attachments[0].contentBytes, "base64"), logo);
  assert.deepEqual(Buffer.from(result.attachments[1].contentBytes, "base64"), image);
  assert.equal(result.from, undefined);
});

test("places Careers at Bond once at the end on the scheduled Brisbane day", () => {
  const html = sourceMail({
    extra: `<tr><td><a href="${careersPage}?from=share">Careers at Bond</a></td></tr>`,
  });
  const result = draft(html, "2026-11-25");
  assert.match(result.body.content, /People &amp; Places[\s\S]+Careers at Bond/);
  assert.equal((result.body.content.match(/>Careers at Bond</g) || []).length, 1);
  assert.equal(result.attachments.length, 3);
  assert.deepEqual(Buffer.from(result.attachments[2].contentBytes, "base64"), careers);
  assert.doesNotMatch(draft(undefined, "2026-11-25", "Thursday").body.content, />Careers at Bond</);
  assert.match(draft(undefined, "2026-11-26", "Thursday").body.content, />Careers at Bond</);
});

test("preserves Bond Safe Links, excludes unrelated links and never embeds remote images", () => {
  const safeLink =
    "https://aus01.safelinks.protection.outlook.com/?url=" + encodeURIComponent(bondStory);
  const html = sourceMail({
    href: safeLink,
    extra:
      '<tr><td><a href="https://example.org/sites/news/SitePages/Outside.aspx">Outside article</a></td></tr>',
  }).replace("cid:Photo%40Outlook", "https://example.org/pixel.png");
  const result = draft(html);
  assert.ok(result.body.content.includes(safeLink));
  assert.doesNotMatch(result.body.content, /example\.org|Outside article/);
  assert.equal(result.attachments.length, 1);
});

test("rejects non-Bond recipients, unrelated messages and invalid Brisbane dates", () => {
  assert.throws(
    () => format(workbook, sourceMail(), "[]", "person@example.org", "2026-11-26"),
    /Bond staff address/,
  );
  assert.throws(
    () => draft(sourceMail({ notice: "A regular message." })),
    /SharePoint digest footer/,
  );
  assert.throws(
    () =>
      draft(sourceMail({ href: "https://outside.sharepoint.com/sites/news/SitePages/Other.aspx" })),
    /No Bond SharePoint news/,
  );
  assert.throws(() => draft(sourceMail(), "2026-02-30"), /Brisbane date/);
});

test("the cloud flow only uses Bond mailbox actions and runs the mark-read action after sending", () => {
  const outlook = "/providers/Microsoft.PowerApps/apis/shared_office365";
  const excel = "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness";
  const trigger = flow.triggers.On_new_email;
  const {
    List_attachments,
    Format_digest,
    Create_draft,
    Send_draft,
    Mark_source_read,
    ...unexpected
  } = flow.actions.If_unread_digest.actions;
  assert.deepEqual(unexpected, {});
  assert.equal(trigger.inputs.host.apiId, outlook);
  assert.equal(trigger.inputs.parameters.from, "@parameters('Mailbox')");
  assert.equal(trigger.inputs.parameters.to, "@parameters('Mailbox')");
  assert.equal(trigger.inputs.parameters.subjectFilter, undefined);
  assert.equal(flow.parameters.DigestTitle, undefined);
  const [{ expression }] = trigger.conditions;
  assert.match(expression, /' shared "'/);
  assert.match(expression, /'" with you'/);
  assert.doesNotMatch(expression, /staff news|DigestTitle/i);
  assert.equal(flow.actions.Get_source.inputs.host.apiId, outlook);
  assert.equal(Format_digest.inputs.host.apiId, excel);
  for (const action of [List_attachments, Create_draft, Send_draft, Mark_source_read]) {
    assert.equal(action.type, "OpenApiConnection");
    assert.equal(action.inputs.host.apiId, outlook);
  }
  assert.match(
    List_attachments.inputs.parameters.Uri,
    /https:\/\/graph\.microsoft\.com\/v1\.0\/me\/messages\//,
  );
  assert.equal(Create_draft.inputs.parameters.Uri, "https://graph.microsoft.com/v1.0/me/messages");
  assert.deepEqual(Format_digest.runAfter, { List_attachments: ["Succeeded"] });
  assert.deepEqual(Create_draft.runAfter, { Format_digest: ["Succeeded"] });
  assert.deepEqual(Send_draft.runAfter, { Create_draft: ["Succeeded"] });
  assert.deepEqual(Mark_source_read.runAfter, { Send_draft: ["Succeeded"] });
  assert.equal(Mark_source_read.inputs.parameters.messageId, "@triggerBody()?['id']");
});
