import fs from "node:fs/promises";
import path from "node:path";
import { defineTool, type Tool } from "@github/copilot-sdk";
import { z } from "zod";
import type { DocumentContext, HostApp, ImportOptions, OfficeOp, UiLocale } from "@officecli-agent/protocol";
import type { TaskpaneLink } from "./taskpane-link.js";

export interface OfficeToolContext {
  hostApp: HostApp;
  documentName?: string;
  locale: UiLocale;
  link: TaskpaneLink;
  /** Per-session scratch directory (absolute). */
  workspaceDir: string;
  allowedDirectories: string[];
  permissionTimeoutMs: number;
}

const EXT: Record<HostApp, string> = { Word: "docx", Excel: "xlsx", PowerPoint: "pptx" };
const MAX_RESULT_CHARS = 60_000;

function clip(text: string): string {
  return text.length > MAX_RESULT_CHARS ? text.slice(0, MAX_RESULT_CHARS) + `\n…(truncated, ${text.length} chars total)` : text;
}

function stringify(value: unknown): string {
  if (typeof value === "string") return value;
  if (value === undefined || value === null) return "(done)";
  return JSON.stringify(value, null, 2);
}

function isInside(file: string, dir: string): boolean {
  const rel = path.relative(dir, file);
  return rel !== "" && !rel.startsWith("..") && !path.isAbsolute(rel);
}

function sanitizeFileName(name: string | undefined, fallback: string): string {
  const base = (name ?? fallback).replace(/[\\/:*?"<>|]/g, "_").trim();
  return base.length > 0 ? base : fallback;
}

export function createOfficeTools(ctx: OfficeToolContext): Tool<unknown>[] {
  const ext = EXT[ctx.hostApp];

  /** Run an operation in the add-in and return a clipped text result. */
  const run = async (op: OfficeOp): Promise<string> => clip(stringify(await ctx.link.runOffice(op)));

  async function ensureReadable(filePath: string): Promise<string> {
    const abs = path.resolve(filePath);
    const roots = [ctx.workspaceDir, ...ctx.allowedDirectories];
    if (!roots.some((root) => isInside(abs, path.resolve(root)))) {
      const approved = await ctx.link.askPermission(
        { kind: "read", title: ctx.locale === "ja" ? "許可フォルダ外のファイルを読み取ります" : "Read a file outside the allowed folders", detail: abs },
        ctx.permissionTimeoutMs,
      );
      if (!approved) throw new Error(`Access to ${abs} was denied by the user.`);
    }
    await fs.access(abs);
    return abs;
  }

  // ---- common -------------------------------------------------------------

  const getContext = defineTool("office_get_document_context", {
    description:
      `Summary of the ${ctx.hostApp} document currently open: file name/path, structure, selection and a text preview. ` +
      "Call this first to understand what the user is looking at.",
    parameters: z.object({}),
    skipPermission: true,
    handler: async () => {
      const result = await ctx.link.runOffice<DocumentContext>({ kind: "getContext" });
      return clip(result.summary);
    },
  });

  const exportDoc = defineTool("office_export_document", {
    description:
      `Save a copy of the LIVE ${ctx.hostApp} document (including unsaved changes) as a .${ext} file in the session workspace and return its absolute path. ` +
      "Use the path with the `officecli` tool (view text/outline/issues/screenshot, query, set, add, batch…) for OOXML-level inspection or editing, " +
      "then bring changes back with office_import_file.",
    parameters: z.object({
      fileName: z.string().optional().describe("File name without extension. Defaults to the document name or 'current'."),
    }),
    skipPermission: true,
    handler: async ({ fileName }) => {
      const fallback = ctx.documentName ? ctx.documentName.replace(/\.[^.]+$/, "") : "current";
      const name = sanitizeFileName(fileName, fallback);
      const target = path.join(ctx.workspaceDir, `${name}.${ext}`);
      await fs.mkdir(ctx.workspaceDir, { recursive: true });
      await ctx.link.runOffice({ kind: "exportFile", targetPath: target });
      const stat = await fs.stat(target);
      return `Exported to ${target} (${stat.size} bytes). Snapshot only: later edits to the live document are not reflected.`;
    },
  });

  const importFile = defineTool("office_import_file", {
    description:
      `Insert content from a .${ext} file on disk into the LIVE ${ctx.hostApp} document. ` +
      "Works for any Office file: other documents the user points to, or a workspace copy you edited with officecli. " +
      (ctx.hostApp === "Word"
        ? "Options: location = Start | End | Replace (whole body) | Selection (replace current selection). Default End."
        : ctx.hostApp === "PowerPoint"
          ? "Options: fromSlide/toSlide (1-based range in the source; all when omitted), afterSlide (1-based position in the open deck; end when omitted), keepSourceFormatting."
          : "Options: sheetNames (all sheets when omitted), afterSheet (sheet name; end when omitted). " +
            "Sheets are ADDED; existing sheets are never replaced, so formulas referencing them stay intact."),
    parameters: z.object({
      filePath: z.string().describe(`Absolute path to a .${ext} file.`),
      options: z
        .object({
          location: z.enum(["Start", "End", "Replace", "Selection"]).optional(),
          fromSlide: z.number().int().positive().optional(),
          toSlide: z.number().int().positive().optional(),
          afterSlide: z.number().int().nonnegative().optional(),
          keepSourceFormatting: z.boolean().optional(),
          sheetNames: z.array(z.string()).optional(),
          afterSheet: z.string().optional(),
        })
        .optional(),
    }),
    handler: async ({ filePath, options }) => {
      const abs = await ensureReadable(filePath);
      if (!abs.toLowerCase().endsWith(`.${ext}`)) {
        return {
          textResultForLlm: `office_import_file only accepts .${ext} files for ${ctx.hostApp}. Read the file with officecli ('view <file> text') and write the content with the typed tools instead.`,
          resultType: "failure" as const,
        };
      }
      return run({ kind: "importFile", path: abs, options: options as ImportOptions | undefined });
    },
  });

  const workspaceInfo = defineTool("workspace_info", {
    description: "Return the session workspace directory (scratch folder for exported/edited copies) and the folders on this PC the agent may read without asking.",
    parameters: z.object({}),
    skipPermission: true,
    handler: async () => ({
      hostApp: ctx.hostApp,
      documentName: ctx.documentName ?? null,
      workspaceDir: ctx.workspaceDir,
      allowedDirectories: ctx.allowedDirectories,
      platform: process.platform,
    }),
  });

  const common: Tool<unknown>[] = [getContext, exportDoc, importFile, workspaceInfo] as Tool<unknown>[];

  // ---- Word ---------------------------------------------------------------

  const word: Tool<unknown>[] = [
    defineTool("word_get_selection", {
      description: "Return the currently selected text in Word and its character range.",
      parameters: z.object({}),
      skipPermission: true,
      handler: async () => run({ kind: "word.getSelection" }),
    }),
    defineTool("word_insert_text", {
      description:
        "Insert text into the live Word document. location: Selection (replaces the selection; default) | Start | End. " +
        "Use \\n for paragraph breaks. Optionally apply a paragraph style name (e.g. 'Heading 1', '見出し 1').",
      parameters: z.object({
        text: z.string(),
        location: z.enum(["Selection", "Start", "End"]).default("Selection"),
        style: z.string().optional(),
      }),
      handler: async ({ text, location, style }) => run({ kind: "word.insertText", text, location, style }),
    }),
    defineTool("word_replace_text", {
      description: "Find and replace text in the live Word document (Word's Find; supports replaceAll and matchCase). Returns the number of replacements.",
      parameters: z.object({
        find: z.string(),
        replace: z.string(),
        matchCase: z.boolean().default(false),
        replaceAll: z.boolean().default(true),
      }),
      handler: async (args) => run({ kind: "word.replaceText", ...args }),
    }),
  ] as Tool<unknown>[];

  // ---- Excel --------------------------------------------------------------

  const excel: Tool<unknown>[] = [
    defineTool("excel_list_sheets", {
      description: "List worksheets in the live workbook with their used range addresses.",
      parameters: z.object({}),
      skipPermission: true,
      handler: async () => run({ kind: "excel.listSheets" }),
    }),
    defineTool("excel_get_range", {
      description: "Read values and formulas from a range of the live workbook. address like 'A1:D20' (used range when omitted); sheet defaults to the active sheet.",
      parameters: z.object({ sheet: z.string().optional(), address: z.string().optional() }),
      skipPermission: true,
      handler: async (args) => run({ kind: "excel.getRange", ...args }),
    }),
    defineTool("excel_set_range", {
      description:
        "Write to the live workbook (undoable). Provide values (2D array matching the address size) and/or formulas (2D array of strings starting with '='). " +
        "Optionally set numberFormat (e.g. '#,##0', 'yyyy-mm-dd'). Use this to apply results computed with officecli to existing cells so references stay intact.",
      parameters: z.object({
        sheet: z.string().optional(),
        address: z.string(),
        values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).optional(),
        formulas: z.array(z.array(z.string())).optional(),
        numberFormat: z.string().optional(),
      }),
      handler: async (args) => run({ kind: "excel.setRange", ...args }),
    }),
  ] as Tool<unknown>[];

  // ---- PowerPoint ---------------------------------------------------------

  const powerpoint: Tool<unknown>[] = [
    defineTool("ppt_set_shape_text", {
      description: "Set the text of a shape on a slide of the live presentation. slide is 1-based; shape is the shape name from office_get_document_context.",
      parameters: z.object({ slide: z.number().int().positive(), shape: z.string(), text: z.string() }),
      handler: async (args) => run({ kind: "ppt.setShapeText", ...args }),
    }),
    defineTool("ppt_add_slide", {
      description: "Add a slide to the live presentation using a layout index of the current master (1-based; 2 is usually Title+Content). Optionally fill title/body placeholders. after: 1-based position (end when omitted).",
      parameters: z.object({
        layout: z.number().int().positive().default(2),
        title: z.string().optional(),
        body: z.string().optional(),
        after: z.number().int().nonnegative().optional(),
      }),
      handler: async (args) => run({ kind: "ppt.addSlide", ...args }),
    }),
    defineTool("ppt_delete_slides", {
      description: "Delete slides by 1-based index (e.g. originals superseded by an import).",
      parameters: z.object({ slides: z.array(z.number().int().positive()).min(1) }),
      handler: async ({ slides }) => run({ kind: "ppt.deleteSlides", slides }),
    }),
  ] as Tool<unknown>[];

  const byApp: Record<HostApp, Tool<unknown>[]> = { Word: word, Excel: excel, PowerPoint: powerpoint };
  return [...common, ...byApp[ctx.hostApp]];
}
