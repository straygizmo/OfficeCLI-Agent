/* global Office, Word, Excel, PowerPoint */
import type {
  DocumentContext,
  ExcelImportOptions,
  HostApp,
  OfficeOp,
  PowerPointImportOptions,
  WordImportOptions,
} from "@officecli-agent/protocol";

export function detectHostApp(): HostApp {
  const host = Office.context.host;
  if (host === Office.HostType.Word) return "Word";
  if (host === Office.HostType.Excel) return "Excel";
  if (host === Office.HostType.PowerPoint) return "PowerPoint";
  throw new Error(`Unsupported Office host: ${String(host)}`);
}

export function detectApiSets(): Record<string, boolean> {
  const sets: Array<[string, string]> = [
    ["WordApi", "1.3"],
    ["WordApi", "1.5"],
    ["ExcelApi", "1.13"],
    ["ExcelApi", "1.16"],
    ["PowerPointApi", "1.2"],
    ["PowerPointApi", "1.5"],
  ];
  const out: Record<string, boolean> = {};
  for (const [name, version] of sets) {
    try {
      out[`${name} ${version}`] = Office.context.requirements.isSetSupported(name, version);
    } catch {
      out[`${name} ${version}`] = false;
    }
  }
  return out;
}

export async function getDocumentName(): Promise<string | undefined> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getFilePropertiesAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded && res.value?.url) {
          const url: string = res.value.url;
          resolve(url.split(/[\\/]/).pop());
        } else resolve(undefined);
      });
    } catch {
      resolve(undefined);
    }
  });
}

/** Dispatch an operation requested by the host. */
export async function runOfficeOp(op: OfficeOp, hostApp: HostApp): Promise<unknown> {
  switch (op.kind) {
    case "execute":
      return executeCode(op.code, hostApp);
    case "exportFile":
      return exportFileBase64();
    case "importFile":
      return importFile(op.base64, op.fileName, op.options, hostApp);
    case "getContext":
      return getContext(hostApp);
  }
}

// ---------------------------------------------------------------------------
// Generic Office.js execution
// ---------------------------------------------------------------------------

type AnyRun = (batch: (context: OfficeExtension.ClientRequestContext) => Promise<unknown>) => Promise<unknown>;

function runnerFor(hostApp: HostApp): AnyRun {
  switch (hostApp) {
    case "Word":
      return (batch) => Word.run((ctx) => batch(ctx));
    case "Excel":
      return (batch) => Excel.run((ctx) => batch(ctx));
    case "PowerPoint":
      return (batch) => PowerPoint.run((ctx) => batch(ctx));
  }
}

async function executeCode(code: string, hostApp: HostApp): Promise<unknown> {
  // eslint-disable-next-line @typescript-eslint/no-implied-eval
  const fn = new Function("context", "hostApp", `"use strict"; return (async () => {\n${code}\n})();`) as (
    context: OfficeExtension.ClientRequestContext,
    hostApp: HostApp,
  ) => Promise<unknown>;
  try {
    return await runnerFor(hostApp)(async (context) => {
      const value = await fn(context, hostApp);
      await context.sync();
      return toSerializable(value);
    });
  } catch (err) {
    throw new Error(describeOfficeError(err));
  }
}

function toSerializable(value: unknown): unknown {
  if (value === undefined || value === null) return value;
  if (typeof value !== "object") return value;
  // Office.js proxy objects expose toJSON() with the loaded properties.
  const maybe = value as { toJSON?: () => unknown };
  if (typeof maybe.toJSON === "function") return maybe.toJSON();
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return String(value);
  }
}

export function describeOfficeError(err: unknown): string {
  if (err instanceof OfficeExtension.Error) {
    const info = err.debugInfo ? ` ${JSON.stringify(err.debugInfo)}` : "";
    return `${err.code}: ${err.message}${info}`;
  }
  if (err instanceof Error) return `${err.name}: ${err.message}`;
  return String(err);
}

// ---------------------------------------------------------------------------
// Export the live document as the compressed (OOXML) file
// ---------------------------------------------------------------------------

export function exportFileBase64(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(`getFileAsync failed: ${res.error?.message ?? "unknown error"}`));
        return;
      }
      const file = res.value;
      const chunks: Uint8Array[] = [];
      let index = 0;
      const finish = (error?: Error) => {
        file.closeAsync(() => {
          if (error) reject(error);
          else resolve(bytesToBase64(concat(chunks)));
        });
      };
      const next = () => {
        if (index >= file.sliceCount) return finish();
        file.getSliceAsync(index, (sliceRes) => {
          if (sliceRes.status !== Office.AsyncResultStatus.Succeeded) {
            finish(new Error(`getSliceAsync(${index}) failed: ${sliceRes.error?.message ?? "unknown error"}`));
            return;
          }
          chunks.push(Uint8Array.from(sliceRes.value.data as number[] | Uint8Array));
          index += 1;
          next();
        });
      };
      next();
    });
  });
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) {
    binary += String.fromCharCode(...bytes.subarray(i, i + step));
  }
  return btoa(binary);
}

// ---------------------------------------------------------------------------
// Import a file into the live document
// ---------------------------------------------------------------------------

async function importFile(base64: string, fileName: string, options: unknown, hostApp: HostApp): Promise<unknown> {
  try {
    switch (hostApp) {
      case "Word": {
        const o = (options ?? {}) as WordImportOptions;
        const location = o.location ?? "End";
        await Word.run(async (context) => {
          if (location === "Selection") {
            context.document.getSelection().insertFileFromBase64(base64, Word.InsertLocation.replace);
          } else {
            const loc =
              location === "Start" ? Word.InsertLocation.start : location === "Replace" ? Word.InsertLocation.replace : Word.InsertLocation.end;
            context.document.body.insertFileFromBase64(base64, loc);
          }
          await context.sync();
        });
        return `Inserted ${fileName} into the document (${location}).`;
      }
      case "PowerPoint": {
        const o = (options ?? {}) as PowerPointImportOptions;
        let before = 0;
        let after = 0;
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          before = slides.items.length;
          context.presentation.insertSlidesFromBase64(base64, {
            formatting: o.formatting === "KeepSourceFormatting" ? PowerPoint.InsertSlideFormatting.keepSourceFormatting : PowerPoint.InsertSlideFormatting.useDestinationTheme,
            sourceSlideIds: o.sourceSlideIds,
            targetSlideId: o.targetSlideId,
          });
          await context.sync();
          slides.load("items");
          await context.sync();
          after = slides.items.length;
        });
        return `Inserted ${after - before} slide(s) from ${fileName}. Slide count: ${before} -> ${after}.`;
      }
      case "Excel": {
        const o = (options ?? {}) as ExcelImportOptions;
        let ids: string[] = [];
        await Excel.run(async (context) => {
          const positionType =
            o.positionType === "Start"
              ? Excel.WorksheetPositionType.beginning
              : o.positionType === "Before"
                ? Excel.WorksheetPositionType.before
                : o.positionType === "After"
                  ? Excel.WorksheetPositionType.after
                  : Excel.WorksheetPositionType.end;
          const result = context.workbook.insertWorksheetsFromBase64(base64, {
            sheetNamesToInsert: o.sheetNames,
            positionType,
            relativeTo: o.relativeTo,
          });
          await context.sync();
          ids = result.value;
        });
        return `Inserted ${ids.length} worksheet(s) from ${fileName}. Worksheet ids: ${ids.join(", ")}`;
      }
    }
  } catch (err) {
    throw new Error(describeOfficeError(err));
  }
}

// ---------------------------------------------------------------------------
// Document context summary
// ---------------------------------------------------------------------------

const PREVIEW_CHARS = 6000;

async function getContext(hostApp: HostApp): Promise<DocumentContext> {
  const documentName = await getDocumentName();
  let summary = "";
  try {
    switch (hostApp) {
      case "Word":
        summary = await Word.run(async (context) => {
          const body = context.document.body;
          const paragraphs = body.paragraphs;
          const selection = context.document.getSelection();
          body.load("text");
          paragraphs.load("items/style,items/text,items/styleBuiltIn");
          selection.load("text");
          await context.sync();
          const headings = paragraphs.items
            .filter((p) => /heading|title/i.test(p.style))
            .slice(0, 60)
            .map((p) => `  [${p.style}] ${p.text.slice(0, 120)}`);
          return [
            `Paragraphs: ${paragraphs.items.length}, characters: ${body.text.length}`,
            `Selection: ${JSON.stringify(selection.text.slice(0, 500))}`,
            headings.length ? `Headings:\n${headings.join("\n")}` : "Headings: (none)",
            `Text preview:\n${body.text.slice(0, PREVIEW_CHARS)}`,
          ].join("\n");
        });
        break;
      case "Excel":
        summary = await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          const active = sheets.getActiveWorksheet();
          const used = active.getUsedRangeOrNullObject(true);
          const selected = context.workbook.getSelectedRange();
          sheets.load("items/name,items/visibility");
          active.load("name");
          used.load("address,rowCount,columnCount,values");
          selected.load("address");
          await context.sync();
          const lines = [`Sheets: ${sheets.items.map((s) => s.name + (s.visibility !== "Visible" ? ` (${s.visibility})` : "")).join(", ")}`, `Active sheet: ${active.name}`, `Selection: ${selected.address}`];
          if (used.isNullObject) lines.push("Used range: (empty)");
          else {
            lines.push(`Used range: ${used.address} (${used.rowCount} x ${used.columnCount})`);
            const rows = used.values.slice(0, 40).map((r) => r.map((v) => String(v ?? "").slice(0, 30)).join("\t"));
            lines.push(`First rows:\n${rows.join("\n")}`);
          }
          return lines.join("\n");
        });
        break;
      case "PowerPoint":
        summary = await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items/id");
          await context.sync();
          const parts: string[] = [`Slides: ${slides.items.length}`];
          const limit = Math.min(slides.items.length, 40);
          for (let i = 0; i < limit; i++) {
            slides.items[i].shapes.load("items/name,items/type,items/id");
          }
          await context.sync();
          const textFrames: Array<{ idx: number; name: string; frame: PowerPoint.TextFrame }> = [];
          for (let i = 0; i < limit; i++) {
            for (const shape of slides.items[i].shapes.items) {
              if (shape.type === "GeometricShape" || shape.type === "Placeholder" || shape.type === "TextBox") {
                const frame = shape.textFrame;
                frame.load("hasText,textRange/text");
                textFrames.push({ idx: i, name: shape.name, frame });
              }
            }
          }
          await context.sync();
          for (let i = 0; i < limit; i++) {
            const texts = textFrames
              .filter((t) => t.idx === i && t.frame.hasText)
              .map((t) => `${t.name}: ${t.frame.textRange.text.replace(/\s+/g, " ").slice(0, 100)}`);
            parts.push(`Slide ${i + 1} (id ${slides.items[i].id}): ${texts.join(" | ") || "(no text)"}`);
          }
          return parts.join("\n");
        });
        break;
    }
  } catch (err) {
    summary = `Could not read document details: ${describeOfficeError(err)}`;
  }
  return {
    hostApp,
    documentName,
    summary: `Host: ${hostApp}\nDocument: ${documentName ?? "(unsaved)"}\n${summary}`,
  };
}
