import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { COPILOT_PROVIDER, DEFAULT_PORT, type ByokProviderInput, type ProviderSummary } from "@officecli-agent/protocol";

/** A BYOK provider as stored in config.json (includes the secret). */
export interface StoredProvider extends ByokProviderInput {
  /** Extra HTTP headers for the provider. */
  headers?: Record<string, string>;
}

export interface HostConfig {
  /** HTTPS/WSS port the task pane connects to. */
  port: number;
  /** Default provider: "copilot" or the name of an entry in `providers`. */
  defaultProvider: string;
  /** BYOK providers selectable from the task pane. */
  providers: StoredProvider[];
  /** Default Copilot model id (CLI default when undefined). */
  model?: string;
  /** Directories the agent may read/write on this PC, besides the workspace. */
  allowedDirectories: string[];
  /** Path to the officecli executable. */
  officecliPath: string;
  /** Path to the GitHub Copilot CLI executable (PATH lookup when undefined). */
  copilotCliPath?: string;
  /**
   * Serve over HTTPS/WSS using localhost.crt / localhost.key from `certDir`.
   * The VSTO add-in talks to the host over plain ws://127.0.0.1, so this is off by default.
   */
  tls: boolean;
  /** Folder containing localhost.crt / localhost.key (office-addin-dev-certs layout). */
  certDir: string;
  /** Root folder for per-session scratch files (exported documents etc.). */
  workspaceRoot: string;
  /** Auto-approve read-only tool calls (file reads, read-only MCP calls). */
  autoApproveReadOnly: boolean;
  /** Seconds to wait for the user to answer a permission prompt before denying. */
  permissionTimeoutSeconds: number;
  /** Also serve the built chat UI (packages/ui/dist) over HTTP for browser-based debugging. */
  serveTaskpane: boolean;
  /** Exit when no add-in has been connected for this many seconds (0 = run until killed). */
  idleExitSeconds: number;
  /** Copilot CLI MCP servers switched off from the task pane (names from mcp-config.json / plugins). */
  disabledMcpServers: string[];
  /** Skills switched off from the task pane (names as the runtime reports them). */
  disabledSkills: string[];
}

export const CONFIG_DIR = path.join(os.homedir(), ".officecli-agent");
export const CONFIG_FILE = path.join(CONFIG_DIR, "config.json");

const defaults: HostConfig = {
  port: DEFAULT_PORT,
  defaultProvider: COPILOT_PROVIDER,
  providers: [],
  allowedDirectories: [],
  officecliPath: "officecli",
  tls: false,
  certDir: path.join(os.homedir(), ".office-addin-dev-certs"),
  workspaceRoot: path.join(CONFIG_DIR, "workspace"),
  autoApproveReadOnly: true,
  permissionTimeoutSeconds: 180,
  serveTaskpane: false,
  idleExitSeconds: 0,
  disabledMcpServers: [],
  disabledSkills: [],
};

/** Command-line overrides: --port <n>, --tls / --no-tls, --serve-ui, --idle-exit <seconds>. */
export function parseArgs(argv: string[]): Partial<HostConfig> {
  const out: Partial<HostConfig> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--port" && argv[i + 1]) out.port = Number(argv[++i]);
    else if (a.startsWith("--port=")) out.port = Number(a.slice(7));
    else if (a === "--tls") out.tls = true;
    else if (a === "--no-tls") out.tls = false;
    else if (a === "--serve-ui") out.serveTaskpane = true;
    else if (a === "--idle-exit" && argv[i + 1]) out.idleExitSeconds = Number(argv[++i]);
    else if (a.startsWith("--idle-exit=")) out.idleExitSeconds = Number(a.slice(12));
  }
  return out;
}

export function loadConfig(overrides: Partial<HostConfig> = {}): HostConfig {
  let fileConfig: Partial<HostConfig> = {};
  if (fs.existsSync(CONFIG_FILE)) {
    try {
      fileConfig = JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8")) as Partial<HostConfig>;
    } catch (err) {
      console.warn(`[config] failed to parse ${CONFIG_FILE}: ${(err as Error).message}`);
    }
  } else {
    fs.mkdirSync(CONFIG_DIR, { recursive: true });
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(defaults, null, 2) + "\n");
  }
  const merged: HostConfig = { ...defaults, ...fileConfig, ...overrides };
  merged.allowedDirectories = (merged.allowedDirectories ?? []).map(expandHome);
  merged.certDir = expandHome(merged.certDir);
  merged.workspaceRoot = expandHome(merged.workspaceRoot);
  merged.providers = Array.isArray(merged.providers) ? merged.providers : [];
  merged.disabledMcpServers = Array.isArray(merged.disabledMcpServers) ? merged.disabledMcpServers : [];
  merged.disabledSkills = Array.isArray(merged.disabledSkills) ? merged.disabledSkills : [];
  fs.mkdirSync(merged.workspaceRoot, { recursive: true });
  return merged;
}

/** Persist the editable parts of the config (providers, default provider). */
export function saveConfig(config: HostConfig): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true });
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2) + "\n", { mode: 0o600 });
}

function fromEnv(name: string | undefined): string | undefined {
  return (name && process.env[name]?.trim()) || undefined;
}

/** A key entered in the pane wins; otherwise the `apiKeyEnv` variable. */
export function resolveApiKey(provider: StoredProvider): string | undefined {
  return provider.apiKey || fromEnv(provider.apiKeyEnv);
}

/** True when the base URL is unset or still the Azure preset's `<resource>` placeholder. */
function baseUrlMissing(baseUrl: string | undefined): boolean {
  return !baseUrl?.trim() || baseUrl.includes("<resource>");
}

/**
 * A base URL entered in the pane wins; for Azure without one, it is built from the
 * resource name in `resourceNameEnv` (https://<resource>.openai.azure.com).
 */
export function resolveBaseUrl(provider: StoredProvider): string | undefined {
  if (!baseUrlMissing(provider.baseUrl)) return provider.baseUrl.trim();
  if (provider.type !== "azure") return undefined;
  const resource = fromEnv(provider.resourceNameEnv);
  return resource ? `https://${resource}.openai.azure.com` : undefined;
}

export function summarizeProviders(config: HostConfig): ProviderSummary[] {
  return [
    { name: COPILOT_PROVIDER, type: "copilot", model: config.model, hasApiKey: true },
    ...config.providers.map((p): ProviderSummary => ({
      name: p.name,
      type: p.type,
      baseUrl: p.baseUrl,
      model: p.model,
      azureApiVersion: p.azureApiVersion,
      apiKeyEnv: p.apiKeyEnv,
      resourceNameEnv: p.resourceNameEnv,
      hasApiKey: Boolean(resolveApiKey(p)),
      apiKeySource: p.apiKey ? "stored" : fromEnv(p.apiKeyEnv) ? "env" : undefined,
    })),
  ];
}

/** Add or update a BYOK provider from task-pane input. An omitted apiKey / *Env keeps the stored one. */
export function upsertProvider(config: HostConfig, input: ByokProviderInput): void {
  const name = input.name.trim();
  if (!name) throw new Error("Provider name is required.");
  if (name.toLowerCase() === COPILOT_PROVIDER) throw new Error(`'${COPILOT_PROVIDER}' is reserved.`);
  const existing = config.providers.find((p) => p.name === name);
  const baseUrl = input.baseUrl.trim();
  const resourceNameEnv = input.type === "azure" ? (input.resourceNameEnv ?? existing?.resourceNameEnv)?.trim() : "";
  if (baseUrl ? !/^https?:\/\//i.test(baseUrl) : !resourceNameEnv) {
    throw new Error("baseUrl must start with http:// or https:// (Azure: or set the resource name environment variable)");
  }

  const next: StoredProvider = {
    ...(existing ?? {}),
    name,
    type: input.type,
    baseUrl,
    model: input.model?.trim() || undefined,
    wireApi: input.wireApi,
    azureApiVersion: input.azureApiVersion?.trim() || undefined,
    modelProvider: input.modelProvider,
  };
  if (input.apiKey !== undefined) next.apiKey = input.apiKey.trim() || undefined;
  if (input.apiKeyEnv !== undefined) next.apiKeyEnv = input.apiKeyEnv.trim() || undefined;
  next.resourceNameEnv = resourceNameEnv || undefined;

  if (existing) Object.assign(existing, next);
  else config.providers.push(next);
  saveConfig(config);
}

export function deleteProvider(config: HostConfig, name: string): void {
  config.providers = config.providers.filter((p) => p.name !== name);
  if (config.defaultProvider === name) config.defaultProvider = COPILOT_PROVIDER;
  saveConfig(config);
}

/** `list` with `name` present or absent, without duplicates, order kept. */
export function withName(list: string[], name: string, present: boolean): string[] {
  const without = list.filter((n) => n !== name);
  return present ? (list.includes(name) ? list : [...list, name]) : without;
}

export function setMcpServerEnabled(config: HostConfig, name: string, enabled: boolean): void {
  config.disabledMcpServers = withName(config.disabledMcpServers, name, !enabled);
  saveConfig(config);
}

export function setSkillEnabled(config: HostConfig, name: string, enabled: boolean): void {
  config.disabledSkills = withName(config.disabledSkills, name, !enabled);
  saveConfig(config);
}

/** The Copilot CLI's configuration directory (never written by this host). */
export function copilotHome(): string {
  return process.env.COPILOT_HOME || path.join(os.homedir(), ".copilot");
}

export function copilotMcpConfigPath(): string {
  return path.join(copilotHome(), "mcp-config.json");
}

export function copilotSkillsDir(): string {
  return path.join(copilotHome(), "skills");
}

export function expandHome(p: string): string {
  if (p.startsWith("~")) return path.join(os.homedir(), p.slice(1));
  return p;
}
