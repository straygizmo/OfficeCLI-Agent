import { test } from "node:test";
import assert from "node:assert/strict";
import { applyToggleToSessions } from "../src/toggles.js";

test("applyToggleToSessions flips every live session, not only the one that asked", async () => {
  const calls: string[] = [];
  const make = (id: string) => ({
    setMcpEnabled: async (name: string, enabled: boolean) => { calls.push(`${id}:mcp:${name}:${enabled}`); },
    setSkillEnabled: async (name: string, enabled: boolean) => { calls.push(`${id}:skill:${name}:${enabled}`); },
  });
  const sessions = new Set([make("word"), make("excel")]);
  await applyToggleToSessions(sessions, "mcp", "gh", false);
  await applyToggleToSessions(sessions, "skill", "officecli", true);
  assert.deepEqual(calls, ["word:mcp:gh:false", "excel:mcp:gh:false", "word:skill:officecli:true", "excel:skill:officecli:true"]);
});

test("applyToggleToSessions keeps going when one session throws", async () => {
  const calls: string[] = [];
  const bad = { setMcpEnabled: async () => { throw new Error("gone"); }, setSkillEnabled: async () => {} };
  const good = { setMcpEnabled: async (n: string) => { calls.push(n); }, setSkillEnabled: async () => {} };
  await applyToggleToSessions([bad, good], "mcp", "gh", true);
  assert.deepEqual(calls, ["gh"]);
});
