/**
 * Wire protocol between the Office task pane (browser) and the local host
 * process (Node). Transported as JSON text frames over a WebSocket.
 */

export type HostApp = "Word" | "Excel" | "PowerPoint";

export const PROTOCOL_VERSION = 1;
export const DEFAULT_PORT = 4320;
export const WS_PATH = "/ws";

// ---------------------------------------------------------------------------
// Office operations: the host asks the add-in (VSTO, C#) to run these against
// the live document through the Office COM object model. The add-in is the
// only place the document is reachable.
// ---------------------------------------------------------------------------

export interface WordImportOptions {
  /** Where to put the imported file's content. Default "End". */
  location?: "Start" | "End" | "Replace" | "Selection";
}

export interface PowerPointImportOptions {
  /** 1-based slide range inside the source file (all slides when omitted). */
  fromSlide?: number;
  toSlide?: number;
  /** Insert after this 1-based slide index (end of deck when omitted). */
  afterSlide?: number;
  /** Keep the source formatting instead of the destination theme. Default false. */
  keepSourceFormatting?: boolean;
}

export interface ExcelImportOptions {
  /** Sheet names inside the source workbook to import (all when omitted). */
  sheetNames?: string[];
  /** Insert after this sheet name (end when omitted). */
  afterSheet?: string;
}

export type ImportOptions = WordImportOptions | PowerPointImportOptions | ExcelImportOptions;

/** Where and how big to place a picture. Fields that do not apply to the host app are ignored. */
export interface InsertImageOptions {
  /** Word: Selection (replaces a non-empty selection; default) | Start | End. */
  location?: "Selection" | "Start" | "End";
  /** Excel: worksheet name (active sheet when omitted). */
  sheet?: string;
  /** Excel: cell for the picture's top-left corner, e.g. "B3" (active cell when omitted). */
  cell?: string;
  /** PowerPoint: 1-based slide (current slide when omitted). */
  slide?: number;
  /** PowerPoint: position in points (centered on the slide when omitted). */
  left?: number;
  top?: number;
  /** Size in points (1 cm ≈ 28.35 pt). One side keeps the aspect ratio; both stretch. Original size when omitted. */
  width?: number;
  height?: number;
  /** Alternative text for accessibility. */
  altText?: string;
}

export type OfficeOp =
  // ---- common ----
  | { kind: "getContext" }
  /**
   * Save a copy of the live document to `targetPath` (the add-in writes the file).
   * `documentName` picks another document open in the same application instead of the active one.
   */
  | { kind: "exportFile"; targetPath: string; documentName?: string }
  /** List the documents open in this application instance (see OpenDocumentSummary). */
  | { kind: "listOpenDocuments" }
  | { kind: "importFile"; path: string; options?: ImportOptions }
  /** Insert a picture file (embedded, not linked) into the live document. */
  | { kind: "insertImage"; path: string; options?: InsertImageOptions }
  // ---- Word ----
  | { kind: "word.getSelection" }
  | { kind: "word.insertText"; text: string; location: "Selection" | "Start" | "End"; style?: string }
  | { kind: "word.replaceText"; find: string; replace: string; matchCase?: boolean; replaceAll?: boolean }
  // ---- Excel ----
  | { kind: "excel.listSheets" }
  | { kind: "excel.getRange"; sheet?: string; address?: string }
  | { kind: "excel.setRange"; sheet?: string; address: string; values?: unknown[][]; formulas?: string[][]; numberFormat?: string }
  // ---- PowerPoint ----
  | { kind: "ppt.setShapeText"; slide: number; shape: string; text: string }
  | { kind: "ppt.addSlide"; layout?: number; title?: string; body?: string; after?: number }
  | { kind: "ppt.deleteSlides"; slides: number[] };

export interface DocumentContext {
  hostApp: HostApp;
  documentName?: string;
  /** Full path of the document on disk, if saved. */
  fullName?: string;
  summary: string;
}

/** One entry of a `listOpenDocuments` result. */
export interface OpenDocumentSummary {
  name: string;
  /** Full path on disk; null for a document that was never saved. */
  path: string | null;
  saved: boolean;
  /** True for the document the task pane is attached to (the one the live tools edit). */
  active: boolean;
}

// ---------------------------------------------------------------------------
// Task pane -> host
// ---------------------------------------------------------------------------

/** Name of the built-in GitHub Copilot provider (uses the Copilot CLI login). */
export const COPILOT_PROVIDER = "copilot";

/** Attachment limits per chat message. */
export const ATTACHMENT_MAX_FILES = 5;
export const ATTACHMENT_MAX_BYTES = 10 * 1024 * 1024;
/** Attachments with these MIME types are shown to the model as images; everything else goes by path. */
export const IMAGE_MIME_TYPES: readonly string[] = ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"];
/** Runtime tool names removed from a session when web search is switched off. */
export const WEB_TOOL_NAMES: readonly string[] = ["web_search", "web_fetch", "github-mcp-server-web_search"];

/** A file attached to a chat message, carried as base64 inside the frame. */
export interface AttachmentInput {
  name: string;
  mimeType: string;
  size: number;
  dataBase64: string;
}

export type ByokProviderType = "openai" | "azure" | "anthropic";

/** A BYOK (bring-your-own-key) model provider, as edited from the task pane. */
export interface ByokProviderInput {
  /** Unique display name, e.g. "OpenAI", "Ollama (local)". */
  name: string;
  type: ByokProviderType;
  /** e.g. https://api.openai.com/v1, http://localhost:11434/v1, https://api.anthropic.com */
  baseUrl: string;
  /** Omit to keep the stored key when updating; empty string clears it. */
  apiKey?: string;
  /** Default model id for this provider (required by the runtime for BYOK). */
  model?: string;
  /** openai/azure only. */
  wireApi?: "completions" | "responses";
  /** Azure API version (azure only). */
  azureApiVersion?: string;
  /** Telemetry hint only, e.g. "ollama". */
  modelProvider?: "openai" | "anthropic" | "azure_openai" | "ollama" | "lm_studio" | "foundry_local" | "llama_cpp";
}

/** What the host tells the task pane about a configured provider (never includes the key). */
export interface ProviderSummary {
  name: string;
  type: ByokProviderType | "copilot";
  baseUrl?: string;
  model?: string;
  hasApiKey: boolean;
}

/** Task pane UI language: Japanese for Japanese environments, English otherwise. */
export type UiLocale = "ja" | "en";

export interface SessionSettings {
  /** Provider name: COPILOT_PROVIDER or the name of a configured BYOK provider. */
  provider?: string;
  model?: string;
  /** Auto-approve read-only tool calls without asking the user. Default true. */
  autoApproveReadOnly?: boolean;
  /**
   * Let the agent list the other documents open in the same application and export copies of them
   * (read-only access; only the active document is edited live). Default false. Applies to new sessions.
   */
  workAcrossOpenFiles?: boolean;
  /** Allow Copilot's web search / URL fetch tools. Default true. Applies to new sessions. */
  webSearch?: boolean;
}

export type ClientMessage =
  | {
      type: "hello";
      protocolVersion: number;
      hostApp: HostApp;
      documentName?: string;
      /** Office build/version string reported by the add-in. */
      officeVersion?: string;
      /** UI language for host-generated text such as permission prompts. Default "en". */
      locale?: UiLocale;
      settings?: SessionSettings;
    }
  | { type: "chat.send"; text: string; attachments?: AttachmentInput[] }
  | { type: "chat.abort" }
  | { type: "office.result"; requestId: string; ok: true; result: unknown }
  | { type: "office.result"; requestId: string; ok: false; error: string }
  | { type: "permission.response"; requestId: string; approved: boolean }
  | { type: "settings.update"; settings: SessionSettings }
  | { type: "session.reset" }
  | { type: "provider.save"; provider: ByokProviderInput }
  | { type: "provider.delete"; name: string }
  /** Switch a Copilot CLI MCP server / skill on or off (live for the running session, persisted for the next). */
  | { type: "mcp.setEnabled"; name: string; enabled: boolean }
  | { type: "skills.setEnabled"; name: string; enabled: boolean };

// ---------------------------------------------------------------------------
// Host -> task pane
// ---------------------------------------------------------------------------

export interface ModelSummary {
  id: string;
  name: string;
}

/** An MCP server known to the Copilot CLI (user config, plugin, built-in), with the pane's on/off state applied. */
export interface McpServerSummary {
  name: string;
  type?: "stdio" | "http" | "sse";
  source: string;
  /** Effective state for new sessions: enabled in the CLI and not switched off from the pane. */
  enabled: boolean;
  /** Disabled in the CLI's own configuration; the pane cannot switch it on. */
  cliDisabled: boolean;
}

export interface SkillSummary {
  name: string;
  description?: string;
  source: string;
  enabled: boolean;
  cliDisabled: boolean;
}

/** What the session loaded from the Copilot CLI configuration, plus where that configuration lives. */
export interface Catalog {
  mcpServers: McpServerSummary[];
  skills: SkillSummary[];
  mcpConfigPath: string;
  skillsDir: string;
}

export type AgentEvent =
  | { kind: "turn.start" }
  | { kind: "message.delta"; messageId: string; text: string }
  | { kind: "message"; messageId: string; text: string }
  | { kind: "reasoning.delta"; text: string }
  | { kind: "tool.start"; toolCallId: string; toolName: string; args?: unknown; mcpServer?: string }
  | { kind: "tool.complete"; toolCallId: string; toolName: string; success: boolean; output?: string; error?: string }
  | { kind: "idle" }
  | { kind: "error"; message: string }
  /** Progress of a provider-hosted tool such as Copilot's web search. */
  | { kind: "server-tool"; name: string; outputIndex: number; status: "in_progress" | "searching" | "completed" };

export interface PermissionSummary {
  requestId: string;
  /** Short label, e.g. "Run shell command". */
  title: string;
  /** Full detail, e.g. the command line or the diff. */
  detail: string;
  kind: string;
}

export type ServerMessage =
  | {
      type: "ready";
      sessionId: string;
      /** Actual host application (the add-in stamps it into `hello`). */
      hostApp: HostApp;
      /** Build timestamp of the running host, for spotting stale processes. */
      hostBuild?: string;
      models: ModelSummary[];
      providers: ProviderSummary[];
      settings: SessionSettings;
      workspaceDir: string;
      allowedDirectories: string[];
      officecliAvailable: boolean;
      catalog: Catalog;
    }
  | { type: "providers"; providers: ProviderSummary[] }
  | { type: "catalog"; catalog: Catalog }
  | { type: "office.request"; requestId: string; op: OfficeOp }
  | { type: "agent.event"; event: AgentEvent }
  | { type: "permission.request"; permission: PermissionSummary }
  | { type: "status"; message: string }
  | { type: "error"; message: string };

export function isClientMessage(value: unknown): value is ClientMessage {
  return typeof value === "object" && value !== null && typeof (value as { type?: unknown }).type === "string";
}
