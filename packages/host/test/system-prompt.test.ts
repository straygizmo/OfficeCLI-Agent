import { test } from "node:test";
import assert from "node:assert/strict";
import { buildSystemPrompt } from "../src/system-prompt.js";

test("system prompt tells the agent where attachments live", () => {
  const text = buildSystemPrompt("Word", "C:\\ws\\abc", true);
  assert.match(text, /Attached files: .*C:\\ws\\abc\\attachments\\/);
  assert.match(text, /office_insert_image/);
});

test("system prompt forbids web access when web search is off", () => {
  const off = buildSystemPrompt("Word", "C:\\ws", true, false, false);
  assert.match(off, /Web search is turned off by the user/);
  assert.match(off, /shell/);
  const on = buildSystemPrompt("Word", "C:\\ws", true, false, true);
  assert.doesNotMatch(on, /Web search is turned off/);
  assert.doesNotMatch(buildSystemPrompt("Word", "C:\\ws", true), /Web search is turned off/);
});
