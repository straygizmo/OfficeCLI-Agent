import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { ATTACHMENT_MAX_BYTES } from "@officecli-agent/protocol";
import { attachmentFooter, isImageMime, sanitizeFileName, saveAttachments, uniquePath, validateAttachments } from "../src/attachments.js";

const mkdtemp = () => fs.mkdtempSync(path.join(os.tmpdir(), "oca-att-"));

test("sanitizeFileName strips directories and reserved characters", () => {
  assert.equal(sanitizeFileName('../..\\evil<>:"|?*.docx'), "evil.docx");
  assert.equal(sanitizeFileName('<>:"|?*'), "attachment");
  assert.equal(sanitizeFileName("  report .docx  "), "report .docx");
});

test("uniquePath numbers duplicates", () => {
  const taken = new Set([path.join("/w", "report.docx"), path.join("/w", "report (2).docx")]);
  assert.equal(uniquePath("/w", "report.docx", (p) => taken.has(p)), path.join("/w", "report (3).docx"));
  assert.equal(uniquePath("/w", "new.png", () => false), path.join("/w", "new.png"));
});

test("isImageMime recognises the supported image types only", () => {
  assert.equal(isImageMime("image/png"), true);
  assert.equal(isImageMime("IMAGE/JPEG"), true);
  assert.equal(isImageMime("application/pdf"), false);
});

test("validateAttachments enforces count and size", () => {
  const one = { name: "a.png", mimeType: "image/png", size: 1, dataBase64: "AA==" };
  assert.equal(validateAttachments([one]), undefined);
  assert.match(validateAttachments(Array(6).fill(one))!, /5/);
  assert.match(validateAttachments([{ ...one, size: ATTACHMENT_MAX_BYTES + 1 }])!, /10 MB/);
});

test("saveAttachments writes under attachments/, judges size by decoded bytes, keeps duplicates apart", async () => {
  const dir = mkdtemp();
  const data = Buffer.from("hello").toString("base64");
  const saved = await saveAttachments(dir, [
    { name: "a.txt", mimeType: "text/plain", size: 999, dataBase64: data },
    { name: "a.txt", mimeType: "text/plain", size: 5, dataBase64: data },
    { name: "pic.png", mimeType: "image/png", size: 5, dataBase64: data },
  ]);
  assert.deepEqual(saved.map((s) => path.basename(s.path)), ["a.txt", "a (2).txt", "pic.png"]);
  assert.equal(saved[0].size, 5);
  assert.equal(saved[2].isImage, true);
  assert.equal(saved[0].isImage, false);
  assert.ok(saved[0].path.startsWith(path.join(dir, "attachments")));
  assert.equal(fs.readFileSync(saved[1].path, "utf8"), "hello");
});

test("saveAttachments refuses a decoded file over the limit", async () => {
  const big = Buffer.alloc(ATTACHMENT_MAX_BYTES + 1).toString("base64");
  await assert.rejects(
    saveAttachments(mkdtemp(), [{ name: "big.bin", mimeType: "application/octet-stream", size: 1, dataBase64: big }]),
    /10 MB/,
  );
});

test("attachmentFooter lists absolute paths and mime types", () => {
  assert.equal(attachmentFooter([]), "");
  const text = attachmentFooter([{ name: "r.docx", path: "C:\\w\\attachments\\r.docx", mimeType: "application/x", size: 1, isImage: false, dataBase64: "" }]);
  assert.match(text, /^\n\nAttached files \(saved on disk\):\n- C:\\w\\attachments\\r\.docx \(application\/x\)$/);
});
