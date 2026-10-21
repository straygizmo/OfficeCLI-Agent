import { test } from "node:test";
import assert from "node:assert/strict";
import { buildSystemPrompt } from "../src/system-prompt.js";

test("system prompt tells the agent where attachments live", () => {
  const text = buildSystemPrompt("Word", "C:\\ws\\abc", true);
  assert.match(text, /Attached files: .*C:\\ws\\abc\\attachments\\/);
  assert.match(text, /office_insert_image/);
});
