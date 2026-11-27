import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";
import test from "node:test";
import { runInNewContext } from "node:vm";

const code = readFileSync(
  new URL("../power-automate/group-post-state.ts", import.meta.url),
  "utf8",
);
const track = runInNewContext(stripTypeScriptTypes(code) + "\nmain;");
const flow = JSON.parse(
  readFileSync(new URL("../power-automate/group-flow-definition.json", import.meta.url), "utf8"),
);

function workbookFixture() {
  const rows = [];
  let hasSheet = false;
  function setCells(row, column, values) {
    for (const [offset, entries] of values.entries()) {
      const target = (rows[row + offset] ??= []);
      for (const [index, value] of entries.entries()) target[column + index] = value;
    }
  }
  const sheet = {
    getUsedRange: () =>
      rows.length
        ? {
            getValues: () => rows.map((row) => [...row]),
            getRowIndex: () => 0,
            getColumnIndex: () => 0,
            getColumnCount: () => rows[0].length,
          }
        : undefined,
    getRange: (address) => {
      assert.equal(address, "A1:C2");
      return { setValues: (values) => setCells(0, 0, values) };
    },
    getRangeByIndexes: (row, column, rowCount, columnCount) => ({
      setValues(values) {
        assert.equal(values.length, rowCount);
        assert.equal(values[0].length, columnCount);
        setCells(row, column, values);
      },
    }),
  };
  const workbook = {
    getWorksheet: () => (hasSheet ? sheet : undefined),
    addWorksheet: (name) => {
      assert.equal(hasSheet, false);
      assert.equal(name, "Digest Bot Group Posts");
      hasSheet = true;
      return sheet;
    },
  };
  return { rows, workbook };
}

test("claims each group post once and records processing after delivery", () => {
  const { rows, workbook } = workbookFixture();
  const firstPost = "+example/group/post-A=";
  const secondPost = "example/group/post-B=";
  assert.equal(track(workbook, firstPost, false), true);
  assert.deepEqual(rows[0], ["Post key", "State", "Updated UTC"]);
  assert.deepEqual(rows[1].slice(0, 2), ["post:" + firstPost, "Pending"]);
  assert.equal(track(workbook, firstPost, false), false);
  assert.equal(track(workbook, secondPost, false), true);
  assert.equal(track(workbook, firstPost, true), true);
  assert.deepEqual(
    rows.map((row) => row[1]),
    ["State", "Processed", "Pending"],
  );
  const processedAt = rows[1][2];
  assert.ok(Number.isFinite(Date.parse(processedAt)));
  assert.equal(track(workbook, firstPost, true), true);
  assert.equal(rows[1][2], processedAt);
  assert.equal(track(workbook, firstPost, false), false);
  assert.equal(rows.length, 3);
});

test("requires an existing claim to finish processing and preserves unrelated worksheet data", () => {
  const { rows, workbook } = workbookFixture();
  assert.throws(() => track(workbook, "unknown-post", true), /never claimed/);
  assert.throws(() => track(workbook, " ", false), /post ID is required/);
  assert.equal(rows.length, 0);
  track(workbook, "valid-post", false);
  assert.throws(() => track(workbook, "unknown-post", true), /never claimed/);
  rows[0][0] = "Unrecognised header";
  assert.throws(() => track(workbook, "another-post", false), /Check the headers/);
  assert.equal(rows.length, 2);
});

test("the group flow sends formatted digests from Bond Staff News to the original sender", () => {
  const groupApi = "/providers/Microsoft.PowerApps/apis/shared_office365groupsmail";
  const groupId = "@parameters('GroupId')";
  const { On_group_email } = flow.triggers;
  const { Get_group_conversation, If_digest_topic } = flow.actions;
  const threadLoop = If_digest_topic.actions.For_each_thread;
  const postLoop = threadLoop.actions.For_each_post;
  const { Get_group_post, Source_sender, If_Bond_digest } = postLoop.actions;
  const { Format_digest, Claim_post, If_claimed } = If_Bond_digest.actions;
  const { Create_draft, Send_draft, Record_processed } = If_claimed.actions;

  assert.deepEqual(Object.keys(flow.triggers), ["On_group_email"]);
  assert.equal(On_group_email.inputs.host.operationId, "OnNewEmailInGroup");
  assert.equal(On_group_email.runtimeConfiguration.concurrency.runs, 1);
  for (const card of [On_group_email, Get_group_conversation, Get_group_post]) {
    assert.equal(card.inputs.host.apiId, groupApi);
    assert.equal(card.inputs.parameters.groupId, groupId);
  }
  assert.match(threadLoop.foreach, /triggerBody\(\)\?\['threads'\]/);
  assert.match(postLoop.foreach, /items\('For_each_thread'\)\?\['posts'\]/);
  assert.equal(Get_group_post.inputs.parameters.threadId, "@items('For_each_thread')?['id']");
  assert.equal(Get_group_post.inputs.parameters.postId, "@items('For_each_post')?['id']");
  assert.match(Source_sender.inputs, /\['sender'\]\?\['emailAddress'\]\?\['address'\]/);
  assert.match(
    If_Bond_digest.expression.and[0].equals[0],
    /endsWith\(outputs\('Source_sender'\), '@bond\.edu\.au'\)/,
  );
  assert.equal(
    Format_digest.inputs.parameters["ScriptParameters/mailbox"],
    "@outputs('Source_sender')",
  );
  assert.match(
    Format_digest.inputs.parameters["ScriptParameters/sourceAttachments"],
    /Get_group_post.*\['attachments'\]/,
  );
  assert.deepEqual(Claim_post.runAfter, { Format_digest: ["Succeeded"] });
  assert.equal(Claim_post.inputs.parameters["ScriptParameters/recordProcessed"], false);
  assert.deepEqual(If_claimed.expression.and[0].equals, ["@body('Claim_post')?['result']", true]);
  assert.deepEqual(Object.keys(If_claimed.else.actions), []);
  assert.equal(Create_draft.inputs.parameters.Uri, "https://graph.microsoft.com/v1.0/me/messages");
  const from = JSON.stringify({ emailAddress: { address: "bondstaffnews@bond.edu.au" } });
  assert.equal(
    Create_draft.inputs.parameters.Body,
    `@string(setProperty(json(body('Format_digest')?['result']), 'from', json('${from}')))`,
  );
  assert.deepEqual(Send_draft.runAfter, { Create_draft: ["Succeeded"] });
  assert.equal(Send_draft.inputs.parameters.messageId, "@body('Create_draft')?['id']");
  assert.deepEqual(Record_processed.runAfter, { Send_draft: ["Succeeded"] });
  assert.equal(Record_processed.inputs.parameters["ScriptParameters/recordProcessed"], true);
  assert.equal(
    Record_processed.inputs.parameters["ScriptParameters/postId"],
    Claim_post.inputs.parameters["ScriptParameters/postId"],
  );
});
