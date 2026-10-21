import { ATTACHMENT_MAX_BYTES, ATTACHMENT_MAX_FILES, IMAGE_MIME_TYPES } from "@officecli-agent/protocol";
import { t } from "./i18n";

/** A file waiting in the composer. `previewUrl` is an object URL for images (revoke it when done). */
export interface PendingAttachment {
  id: string;
  file: File;
  previewUrl?: string;
}

/** What the chat (and local history) keeps about an attachment: never the bytes. */
export interface ShownAttachment {
  name: string;
  mimeType: string;
  size: number;
  /** Small data URL (≤ 64 px) for images. */
  thumbnail?: string;
}

const IMAGE_EXT_TO_MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
};

const EXT_TO_MIME: Record<string, string> = {
  ...IMAGE_EXT_TO_MIME,
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf",
  txt: "text/plain",
  md: "text/markdown",
  csv: "text/csv",
  json: "application/json",
};

/** The browser's type when it knows one, otherwise a guess from the extension (WebView2 leaves `type` empty for some drops). */
export function mimeTypeOf(file: File): string {
  if (file.type) return file.type.toLowerCase();
  const ext = file.name.split(".").pop()?.toLowerCase() ?? "";
  return EXT_TO_MIME[ext] ?? "application/octet-stream";
}

export function isImageFile(file: File): boolean {
  return IMAGE_MIME_TYPES.includes(mimeTypeOf(file));
}

let counter = 0;

/** Add files to the pending list within the limits; returns the new list and an error text when something was refused. */
export function acceptFiles(current: PendingAttachment[], incoming: File[]): { next: PendingAttachment[]; error?: string } {
  const next = [...current];
  let error: string | undefined;
  for (const file of incoming) {
    if (next.length >= ATTACHMENT_MAX_FILES) {
      error = t.tooManyFiles(ATTACHMENT_MAX_FILES);
      break;
    }
    if (file.size > ATTACHMENT_MAX_BYTES) {
      error = t.fileTooLarge(file.name, ATTACHMENT_MAX_BYTES / (1024 * 1024));
      continue;
    }
    next.push({ id: `att-${Date.now()}-${counter++}`, file, previewUrl: isImageFile(file) ? URL.createObjectURL(file) : undefined });
  }
  return { next, error };
}

export function readAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("read failed"));
    reader.onload = () => {
      const url = String(reader.result);
      resolve(url.slice(url.indexOf(",") + 1));
    };
    reader.readAsDataURL(file);
  });
}

/** Small JPEG data URL for an image file; undefined for non-images or when decoding fails. */
export async function makeThumbnail(file: File, maxPx = 64): Promise<string | undefined> {
  if (!isImageFile(file)) return undefined;
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = () => reject(new Error("decode failed"));
      el.src = url;
    });
    const scale = Math.min(1, maxPx / Math.max(img.width, img.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(img.width * scale));
    canvas.height = Math.max(1, Math.round(img.height * scale));
    canvas.getContext("2d")?.drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL("image/jpeg", 0.7);
  } catch {
    return undefined;
  } finally {
    URL.revokeObjectURL(url);
  }
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
