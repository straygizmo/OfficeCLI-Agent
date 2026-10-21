import fs from "node:fs/promises";
import path from "node:path";
import { randomUUID } from "node:crypto";
import type {
  CopilotClient,
  CopilotSession,
  PermissionHandler,
  PermissionRequest,
  ProviderConfig,
  SessionEvent,
} from "@github/copilot-sdk";
import { COPILOT_PROVIDER, WEB_TOOL_NAMES, type AgentEvent, type AttachmentInput, type Catalog, type HostApp, type SessionSettings, type UiLocale } from "@officecli-agent/protocol";
import { saveAttachments, toMessageOptions } from "./attachments.js";
import { buildCatalog } from "./catalog.js";
import { resolveApiKey, type HostConfig } from "./config.js";
import { createOfficeTools } from "./office-tools.js";
import { createOfficecliTool } from "./officecli-tool.js";
import { buildSystemPrompt } from "./system-prompt.js";
import type { TaskpaneLink } from "./taskpane-link.js";

export interface AgentSessionOptions {
  client: CopilotClient;
  config: HostConfig;
  link: TaskpaneLink;
  hostApp: HostApp;
  documentName?: string;
  locale: UiLocale;
  settings: SessionSettings;
  officecliAvailable: boolean;
}

/** One Copilot session bound to one task pane connection. */
export class AgentSession {
  readonly id = randomUUID();
  readonly workspaceDir: string;
  private session: CopilotSession | undefined;
  private unsubscribe: (() => void) | undefined;
  private settings: SessionSettings;
  private catalogValue: Catalog = { mcpServers: [], skills: [], mcpConfigPath: "", skillsDir: "" };

  constructor(private readonly opts: AgentSessionOptions) {
    this.settings = { ...opts.settings };
    this.workspaceDir = path.join(opts.config.workspaceRoot, this.id.slice(0, 8));
  }

  async start(): Promise<void> {
    const { client, config, link, hostApp, documentName, locale, officecliAvailable } = this.opts;
    await fs.mkdir(this.workspaceDir, { recursive: true });

    // Fixed for the lifetime of the session (tools and prompt are registered once); a change applies to the next new chat.
    const workAcrossOpenFiles = this.settings.workAcrossOpenFiles ?? false;
    const tools = createOfficeTools({
      hostApp,
      documentName,
      locale,
      link,
      workspaceDir: this.workspaceDir,
      allowedDirectories: config.allowedDirectories,
      permissionTimeoutMs: config.permissionTimeoutSeconds * 1000,
      workAcrossOpenFiles,
    });

    const { provider, model } = this.resolveProvider();

    if (officecliAvailable) {
      tools.push(
        createOfficecliTool({
          officecliPath: config.officecliPath,
          locale,
          workspaceDir: this.workspaceDir,
          link,
          autoApproveReadOnly: () => this.settings.autoApproveReadOnly ?? config.autoApproveReadOnly,
          permissionTimeoutMs: config.permissionTimeoutSeconds * 1000,
        }),
      );
    }

    this.session = await client.createSession({
      model,
      provider,
      streaming: true,
      workingDirectory: this.workspaceDir,
      additionalDirectories: config.allowedDirectories,
      tools,
      // Load the MCP servers and skills the user's Copilot CLI has (~/.copilot); the pane shows and toggles them.
      enableConfigDiscovery: true,
      enableSkills: true,
      disabledMcpServers: config.disabledMcpServers,
      disabledSkills: config.disabledSkills,
      excludedTools: this.settings.webSearch === false ? [...WEB_TOOL_NAMES] : undefined,
      systemMessage: {
        mode: "append",
        content: buildSystemPrompt(hostApp, this.workspaceDir, officecliAvailable, workAcrossOpenFiles, this.settings.webSearch !== false),
      },
      onPermissionRequest: this.permissionHandler,
    });

    this.unsubscribe = this.session.on((event) => this.forward(event));
    await this.refreshCatalog();
  }

  get sessionId(): string {
    return this.session?.sessionId ?? this.id;
  }

  get link(): TaskpaneLink {
    return this.opts.link;
  }

  /** MCP servers and skills loaded from the Copilot CLI configuration, with the pane's on/off state. */
  get catalog(): Catalog {
    return this.catalogValue;
  }

  async refreshCatalog(): Promise<Catalog> {
    const { client, config } = this.opts;
    const session = this.session;
    this.catalogValue = await buildCatalog(
      () => client.rpc.mcp.discover({}),
      () => (session ? session.rpc.skills.list() : Promise.resolve({ skills: [] })),
      config,
    );
    return this.catalogValue;
  }

  /** Flip an MCP server for the running session; the persisted choice (config.json) is the caller's job. */
  async setMcpEnabled(name: string, enabled: boolean): Promise<void> {
    if (!this.session) return;
    try {
      if (enabled) await this.session.rpc.mcp.enable({ serverName: name });
      else await this.session.rpc.mcp.disable({ serverName: name });
    } catch (err) {
      console.warn(`[session ${this.id}] mcp.${enabled ? "enable" : "disable"}('${name}') failed: ${(err as Error).message}`);
    }
  }

  async setSkillEnabled(name: string, enabled: boolean): Promise<void> {
    if (!this.session) return;
    try {
      if (enabled) await this.session.rpc.skills.enable({ name });
      else await this.session.rpc.skills.disable({ name });
    } catch (err) {
      console.warn(`[session ${this.id}] skills.${enabled ? "enable" : "disable"}('${name}') failed: ${(err as Error).message}`);
    }
  }

  async send(text: string, attachments: AttachmentInput[] = []): Promise<void> {
    if (!this.session) throw new Error("Session not started");
    const saved = attachments.length ? await saveAttachments(this.workspaceDir, attachments) : [];
    await this.session.send(toMessageOptions(text, saved));
  }

  async abort(): Promise<void> {
    await this.session?.abort();
  }

  updateSettings(settings: SessionSettings): void {
    this.settings = { ...this.settings, ...settings };
  }

  async destroy(): Promise<void> {
    this.unsubscribe?.();
    try {
      if (this.session) {
        const id = this.session.sessionId;
        await this.session.disconnect();
        await this.opts.client.deleteSession(id);
      }
    } catch (err) {
      console.warn(`[session ${this.id}] destroy failed: ${(err as Error).message}`);
    }
    this.session = undefined;
  }

  // -------------------------------------------------------------------------

  /** Map the selected provider name to a Copilot SDK ProviderConfig (BYOK) or none (Copilot). */
  private resolveProvider(): { provider: ProviderConfig | undefined; model: string | undefined } {
    const { config } = this.opts;
    const name = this.settings.provider || config.defaultProvider || COPILOT_PROVIDER;
    if (name === COPILOT_PROVIDER) {
      return { provider: undefined, model: this.settings.model || config.model };
    }
    const stored = config.providers.find((p) => p.name === name);
    if (!stored) throw new Error(`Unknown provider '${name}'. Configure it in the task pane settings.`);
    const model = this.settings.model || stored.model;
    if (!model) throw new Error(`Provider '${name}' needs a model id (BYOK requires an explicit model).`);
    const provider: ProviderConfig = {
      type: stored.type,
      baseUrl: stored.baseUrl,
      apiKey: resolveApiKey(stored),
      wireApi: stored.type === "anthropic" ? undefined : (stored.wireApi ?? defaultWireApi(stored.type, stored.baseUrl)),
      modelProvider: stored.modelProvider,
      azure: stored.type === "azure" && stored.azureApiVersion ? { apiVersion: stored.azureApiVersion } : undefined,
      headers: stored.headers,
    };
    return { provider, model };
  }

  private emit(event: AgentEvent): void {
    this.opts.link.send({ type: "agent.event", event });
  }

  private forward(event: SessionEvent): void {
    switch (event.type) {
      case "assistant.turn_start":
        this.emit({ kind: "turn.start" });
        break;
      case "assistant.message_delta":
        this.emit({ kind: "message.delta", messageId: event.data.messageId, text: event.data.deltaContent });
        break;
      case "assistant.message":
        this.emit({ kind: "message", messageId: event.data.messageId, text: event.data.content });
        break;
      case "assistant.reasoning_delta":
        this.emit({ kind: "reasoning.delta", text: event.data.deltaContent });
        break;
      case "tool.execution_start":
        this.emit({
          kind: "tool.start",
          toolCallId: event.data.toolCallId,
          // MCP tools are registered as "<server>-<tool>"; show the bare tool name next to the server.
          toolName: event.data.mcpToolName ?? stripServerPrefix(event.data.toolName, event.data.mcpServerName),
          args: event.data.arguments,
          mcpServer: event.data.mcpServerName,
        });
        break;
      case "tool.execution_complete":
        this.emit({
          kind: "tool.complete",
          toolCallId: event.data.toolCallId,
          toolName: (event.data as { toolName?: string }).toolName ?? "",
          success: event.data.success,
          output: event.data.result?.content?.slice(0, 4000),
          error: event.data.error?.message,
        });
        break;
      case "assistant.server_tool_progress":
        this.emit({
          kind: "server-tool",
          name: event.data.kind,
          outputIndex: event.data.outputIndex,
          status: event.data.status as "in_progress" | "searching" | "completed",
        });
        break;
      case "session.idle":
        this.emit({ kind: "idle" });
        break;
      case "session.error":
        this.emit({ kind: "error", message: event.data.message });
        break;
      default:
        break;
    }
  }

  private permissionHandler: PermissionHandler = async (request) => {
    const { config, link } = this.opts;
    const autoReadOnly = this.settings.autoApproveReadOnly ?? config.autoApproveReadOnly;

    // Handler results use the decision vocabulary ("approve-once" / "reject"), not the
    // recorded-outcome vocabulary ("approved" / "denied-…"); the MCP flow rejects the latter.
    if (request.kind === "custom-tool" || (request.kind === "read" && autoReadOnly)) {
      return { kind: "approve-once" };
    }
    if (request.kind === "mcp" && autoReadOnly && request.readOnly) {
      return { kind: "approve-once" };
    }

    const summary = describePermission(request, this.opts.locale);
    const approved = await link.askPermission(summary, config.permissionTimeoutSeconds * 1000);
    return approved ? { kind: "approve-once", approvedInteractively: true } : { kind: "reject" };
  };
}

/**
 * OpenAI and Azure OpenAI are assumed to serve gpt-5.6 or later, i.e. the Responses API.
 * Other OpenAI-compatible servers (Ollama, vLLM, LM Studio, …) generally only speak chat completions.
 */
function defaultWireApi(type: string, baseUrl: string): "responses" | "completions" {
  return type === "azure" || /api\.openai\.com/i.test(baseUrl) ? "responses" : "completions";
}

function stripServerPrefix(toolName: string, serverName: string | undefined): string {
  if (serverName && toolName.startsWith(serverName + "-")) return toolName.slice(serverName.length + 1);
  return toolName;
}

function describePermission(request: PermissionRequest, locale: UiLocale): { kind: string; title: string; detail: string } {
  const ja = locale === "ja";
  switch (request.kind) {
    case "shell":
      return { kind: request.kind, title: ja ? "シェルコマンドを実行します" : "Run a shell command", detail: `${request.intention}\n\n${request.fullCommandText}` };
    case "write":
      return { kind: request.kind, title: ja ? `ファイルを書き換えます: ${request.fileName}` : `Modify a file: ${request.fileName}`, detail: request.diff || request.intention };
    case "mcp": {
      const tool = stripServerPrefix(request.toolName, request.serverName);
      const label = tool === request.serverName ? request.serverName : `${request.serverName} ${tool}`;
      return { kind: request.kind, title: ja ? `${label} を実行します` : `Run ${label}`, detail: request.args === undefined ? "" : JSON.stringify(request.args, null, 2) };
    }
    case "url":
      return { kind: request.kind, title: ja ? "URLにアクセスします" : "Access a URL", detail: `${request.intention}\n\n${request.url}` };
    case "read":
      return { kind: request.kind, title: ja ? "ファイルを読み取ります" : "Read a file", detail: JSON.stringify(request, null, 2) };
    default:
      return { kind: request.kind, title: ja ? `許可が必要です (${request.kind})` : `Permission required (${request.kind})`, detail: JSON.stringify(request, null, 2).slice(0, 4000) };
  }
}
