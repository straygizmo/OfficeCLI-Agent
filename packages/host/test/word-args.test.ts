import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeReplaceArgs } from "../src/word-args.js";

test("word_replace_text arguments pass through with defaults", () => {
  assert.deepEqual(normalizeReplaceArgs({ find: "a", replace: "b" }), { find: "a", replace: "b", matchCase: false, replaceAll: true });
  assert.deepEqual(normalizeReplaceArgs({ find: "a", replace: "", matchCase: true, replaceAll: false }), { find: "a", replace: "", matchCase: true, replaceAll: false });
});

test("word_replace_text accepts the names models use instead of find/replace", () => {
  // Seen on the real device: { text, replaceWith }.
  assert.deepEqual(normalizeReplaceArgs({ text: "第3四半期 営業報告書", replaceWith: "2026年度 第3四半期 営業報告書", matchCase: true, replaceAll: false }), {
    find: "第3四半期 営業報告書",
    replace: "2026年度 第3四半期 営業報告書",
    matchCase: true,
    replaceAll: false,
  });
  assert.equal(normalizeReplaceArgs({ search: "a", replacement: "b" }).replace, "b");
  assert.equal(normalizeReplaceArgs({ oldText: "a", newText: "b" }).find, "a");
  // The real names win over synonyms.
  assert.deepEqual(normalizeReplaceArgs({ find: "a", text: "x", replace: "b", with: "y" }), { find: "a", replace: "b", matchCase: false, replaceAll: true });
});

test("word_replace_text without anything to find says which parameters to use", () => {
  assert.throws(() => normalizeReplaceArgs({ replace: "b" }), /`find`.*`replace`/);
  assert.throws(() => normalizeReplaceArgs({ find: "" }), /`find`/);
  assert.throws(() => normalizeReplaceArgs(undefined), /`find`/);
});
