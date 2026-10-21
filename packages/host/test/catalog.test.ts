import { test } from "node:test";
import assert from "node:assert/strict";
import { buildCatalog, mergeMcpServers, mergeSkills } from "../src/catalog.js";
import { copilotMcpConfigPath, copilotSkillsDir, type HostConfig } from "../src/config.js";

test("mergeMcpServers applies the pane's disabled list on top of the CLI's", () => {
  const out = mergeMcpServers(
    [
      { name: "gh", type: "http", source: "user", enabled: true },
      { name: "fs", type: "stdio", source: "user", enabled: false },
    ],
    ["gh"],
  );
  assert.deepEqual(out, [
    { name: "fs", type: "stdio", source: "user", enabled: false, cliDisabled: true },
    { name: "gh", type: "http", source: "user", enabled: false, cliDisabled: false },
  ]);
});

test("mergeMcpServers drops an unknown transport type", () => {
  const out = mergeMcpServers([{ name: "x", type: "weird", source: "plugin", enabled: true }], []);
  assert.deepEqual(out, [{ name: "x", type: undefined, source: "plugin", enabled: true, cliDisabled: false }]);
});

test("mergeSkills tells CLI-disabled from pane-disabled", () => {
  const out = mergeSkills(
    [
      { name: "a", description: "A", source: "user", enabled: false },
      { name: "b", description: "B", source: "builtin", enabled: false },
      { name: "c", description: "C", source: "user", enabled: true },
    ],
    ["a"],
  );
  assert.deepEqual(out, [
    { name: "a", description: "A", source: "user", enabled: false, cliDisabled: false },
    { name: "b", description: "B", source: "builtin", enabled: false, cliDisabled: true },
    { name: "c", description: "C", source: "user", enabled: true, cliDisabled: false },
  ]);
});

test("buildCatalog survives a failing source", async () => {
  const cfg = { disabledMcpServers: [], disabledSkills: [] } as unknown as HostConfig;
  const cat = await buildCatalog(
    () => Promise.reject(new Error("no mcp")),
    async () => ({ skills: [{ name: "s", description: "", source: "user", enabled: true }] }),
    cfg,
  );
  assert.deepEqual(cat.mcpServers, []);
  assert.equal(cat.skills[0].name, "s");
  assert.equal(cat.mcpConfigPath, copilotMcpConfigPath());
  assert.equal(cat.skillsDir, copilotSkillsDir());
});
