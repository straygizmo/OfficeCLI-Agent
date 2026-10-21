import { test } from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { copilotHome, copilotMcpConfigPath, copilotSkillsDir, loadConfig, withName } from "../src/config.js";

test("withName adds once and removes", () => {
  assert.deepEqual(withName([], "a", true), ["a"]);
  assert.deepEqual(withName(["a"], "a", true), ["a"]);
  assert.deepEqual(withName(["a", "b"], "a", false), ["b"]);
  assert.deepEqual(withName(["b"], "a", false), ["b"]);
});

test("copilot paths honour COPILOT_HOME", () => {
  const before = process.env.COPILOT_HOME;
  try {
    process.env.COPILOT_HOME = "C:\\ch";
    assert.equal(copilotMcpConfigPath(), path.join("C:\\ch", "mcp-config.json"));
    assert.equal(copilotSkillsDir(), path.join("C:\\ch", "skills"));
    delete process.env.COPILOT_HOME;
    assert.equal(copilotHome(), path.join(os.homedir(), ".copilot"));
  } finally {
    if (before !== undefined) process.env.COPILOT_HOME = before;
  }
});

test("loadConfig defaults the disabled lists to empty arrays", () => {
  const cfg = loadConfig({ disabledMcpServers: undefined as unknown as string[], disabledSkills: undefined as unknown as string[] });
  assert.deepEqual(cfg.disabledMcpServers, []);
  assert.deepEqual(cfg.disabledSkills, []);
});
