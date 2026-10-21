import type { Catalog, McpServerSummary, SkillSummary } from "@officecli-agent/protocol";
import { copilotMcpConfigPath, copilotSkillsDir, type HostConfig } from "./config.js";

/** Shape of one entry from the runtime's `mcp.discover` RPC (the fields we use). */
export interface DiscoveredServer {
  name: string;
  type?: string;
  source: string;
  enabled: boolean;
}

/** Shape of one entry from the session's `skills.list` RPC (the fields we use). */
export interface ListedSkill {
  name: string;
  description?: string;
  source: string;
  enabled: boolean;
}

const TRANSPORTS = new Set(["stdio", "http", "sse"]);

/**
 * The CLI's own disabled state wins: a server the CLI disabled is shown off and cannot be switched
 * on from the pane. A server the CLI has on is shown off only when the pane switched it off.
 */
export function mergeMcpServers(discovered: DiscoveredServer[], disabled: string[]): McpServerSummary[] {
  return discovered
    .map<McpServerSummary>((d) => ({
      name: d.name,
      type: d.type && TRANSPORTS.has(d.type) ? (d.type as McpServerSummary["type"]) : undefined,
      source: d.source,
      enabled: d.enabled && !disabled.includes(d.name),
      cliDisabled: !d.enabled,
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * `skills.list` runs after the session started with our `disabledSkills`, so it reports
 * `enabled: false` for both CLI-disabled and pane-disabled skills; our own list tells them apart.
 */
export function mergeSkills(skills: ListedSkill[], disabled: string[]): SkillSummary[] {
  return skills
    .map<SkillSummary>((s) => ({
      name: s.name,
      description: s.description || undefined,
      source: s.source,
      enabled: s.enabled,
      cliDisabled: !s.enabled && !disabled.includes(s.name),
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/** Build the catalog; a source that throws (older runtime, policy) yields an empty list instead of failing the session. */
export async function buildCatalog(
  discover: () => Promise<{ servers: DiscoveredServer[] }>,
  listSkills: () => Promise<{ skills: ListedSkill[] }>,
  config: HostConfig,
): Promise<Catalog> {
  const [servers, skills] = await Promise.all([
    discover().then((r) => r.servers, warnAndEmpty<DiscoveredServer>("mcp.discover")),
    listSkills().then((r) => r.skills, warnAndEmpty<ListedSkill>("skills.list")),
  ]);
  return {
    mcpServers: mergeMcpServers(servers, config.disabledMcpServers),
    skills: mergeSkills(skills, config.disabledSkills),
    mcpConfigPath: copilotMcpConfigPath(),
    skillsDir: copilotSkillsDir(),
  };
}

function warnAndEmpty<T>(what: string): (err: unknown) => T[] {
  return (err) => {
    console.warn(`[catalog] ${what} failed: ${(err as Error).message ?? String(err)}`);
    return [];
  };
}
