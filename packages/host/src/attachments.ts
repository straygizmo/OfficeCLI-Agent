import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { ATTACHMENT_MAX_BYTES, ATTACHMENT_MAX_FILES, IMAGE_MIME_TYPES, type AttachmentInput } from "@officecli-agent/protocol";

/** An attachment written to disk under the session workspace. */
export interface SavedAttachment {
  name: string;
  /** Absolute path of the saved file. */
  path: string;
  mimeType: string;
  /** Decoded size in bytes. */
  size: number;
  isImage: boolean;
  dataBase64: string;
}

const MAX_MB = ATTACHMENT_MAX_BYTES / (1024 * 1024);

/** Keep only a safe base name: no directories, no Windows-reserved characters, no control characters. */
export function sanitizeFileName(name: string): string {
  const base = name.split(/[\\/]/).pop() ?? "";
  // eslint-disable-next-line no-control-regex
  const cleaned = base.replace(/[<>:"|?*\u0000-\u001f]/g, "").replace(/^[\s.]+|[\s.]+$/g, "");
  return cleaned || "attachment";
}

/** First free path for `fileName` in `dir`: name.ext, name (2).ext, name (3).ext, … */
export function uniquePath(dir: string, fileName: string, exists: (p: string) => boolean): string {
  const ext = path.extname(fileName);
  const stem = fileName.slice(0, fileName.length - ext.length);
  let candidate = path.join(dir, fileName);
  for (let n = 2; exists(candidate); n++) candidate = path.join(dir, `${stem} (${n})${ext}`);
  return candidate;
}

export function isImageMime(mimeType: string): boolean {
  return IMAGE_MIME_TYPES.includes(mimeType.toLowerCase());
}

/** Error text when the batch breaks the limits, otherwise undefined. */
export function validateAttachments(list: AttachmentInput[]): string | undefined {
  if (list.length > ATTACHMENT_MAX_FILES) return `Too many attachments: at most ${ATTACHMENT_MAX_FILES} files per message.`;
  for (const a of list) {
    if (a.size > ATTACHMENT_MAX_BYTES) return `Attachment '${a.name}' is larger than ${MAX_MB} MB.`;
  }
  return undefined;
}

/** Decode and write every attachment into `<workspaceDir>/attachments/`. Throws on a limit violation. */
export async function saveAttachments(workspaceDir: string, list: AttachmentInput[]): Promise<SavedAttachment[]> {
  const error = validateAttachments(list);
  if (error) throw new Error(error);
  const dir = path.join(workspaceDir, "attachments");
  await fs.mkdir(dir, { recursive: true });
  const saved: SavedAttachment[] = [];
  for (const a of list) {
    const buffer = Buffer.from(a.dataBase64, "base64");
    if (buffer.length > ATTACHMENT_MAX_BYTES) throw new Error(`Attachment '${a.name}' is larger than ${MAX_MB} MB.`);
    const name = sanitizeFileName(a.name);
    const target = uniquePath(dir, name, existsSync);
    await fs.writeFile(target, buffer);
    saved.push({ name, path: target, mimeType: a.mimeType, size: buffer.length, isImage: isImageMime(a.mimeType), dataBase64: a.dataBase64 });
  }
  return saved;
}

/** Text appended to the user's prompt so the agent knows where the files are. */
export function attachmentFooter(saved: SavedAttachment[]): string {
  if (saved.length === 0) return "";
  return "\n\nAttached files (saved on disk):\n" + saved.map((s) => `- ${s.path} (${s.mimeType})`).join("\n");
}

export interface BlobAttachment {
  type: "blob";
  data: string;
  mimeType: string;
  displayName?: string;
}

/**
 * Prompt text plus SDK attachments for one message. Documents are never passed to the runtime
 * as file attachments (it would read binary files into the prompt); their paths in the footer
 * are enough for officecli. Images go as blobs so a vision-capable model sees them.
 */
export function toMessageOptions(text: string, saved: SavedAttachment[]): { prompt: string; attachments: BlobAttachment[] | undefined } {
  const blobs = saved.filter((s) => s.isImage).map<BlobAttachment>((s) => ({ type: "blob", data: s.dataBase64, mimeType: s.mimeType, displayName: s.name }));
  return { prompt: text + attachmentFooter(saved), attachments: blobs.length ? blobs : undefined };
}
