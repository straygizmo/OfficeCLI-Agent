/** The part of a session a toggle needs. */
export interface ToggleTarget {
  setMcpEnabled(name: string, enabled: boolean): Promise<void>;
  setSkillEnabled(name: string, enabled: boolean): Promise<void>;
}

/**
 * The on/off choice is shared through config.json, so every connected pane's live session must
 * follow it, not only the pane that flipped the switch. A failing session does not stop the others.
 */
export async function applyToggleToSessions(sessions: Iterable<ToggleTarget>, kind: "mcp" | "skill", name: string, enabled: boolean): Promise<void> {
  for (const s of sessions) {
    try {
      if (kind === "mcp") await s.setMcpEnabled(name, enabled);
      else await s.setSkillEnabled(name, enabled);
    } catch (err) {
      console.warn(`[host] ${kind} toggle '${name}' failed on one session: ${(err as Error).message}`);
    }
  }
}
