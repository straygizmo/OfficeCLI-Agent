import { test } from "node:test";
import assert from "node:assert/strict";
import { isOfficecliOnlyShell, splitPipeline, usesNativeRenderer } from "../src/officecli-shell.js";
import { forceHtmlRenderer } from "../src/officecli-tool.js";

function shell(fullCommandText: string, extra: Record<string, unknown> = {}) {
  const ids = (splitPipeline(fullCommandText) ?? []).map((argv) => ({ identifier: argv[0] ?? "", readOnly: false }));
  return {
    kind: "shell",
    canOfferSessionApproval: false,
    commands: ids,
    fullCommandText,
    hasWriteFileRedirection: false,
    intention: "",
    possiblePaths: [],
    possibleUrls: [],
    ...extra,
  } as unknown as Parameters<typeof isOfficecliOnlyShell>[0];
}

test("plain officecli commands are recognised", () => {
  assert.ok(isOfficecliOnlyShell(shell('officecli set "C:\\ws\\a.docx" /body/p[1] --prop text="Hello; world"')));
  assert.ok(isOfficecliOnlyShell(shell("officecli.exe view C:\\ws\\a.docx outline")));
  assert.ok(isOfficecliOnlyShell(shell('"C:\\Program Files\\OfficeCLI\\officecli.exe" get a.xlsx /Sheet1')));
});

test("a printing command piped into officecli is recognised", () => {
  assert.ok(isOfficecliOnlyShell(shell(`echo '[{"op":"set","path":"/Sheet1/A1","value":"x > y; z"}]' | officecli batch C:\\ws\\d.xlsx --json`)));
  assert.ok(isOfficecliOnlyShell(shell("Get-Content C:\\ws\\ops.json | officecli batch C:\\ws\\d.xlsx --json")));
});

test("anything beyond officecli needs approval", () => {
  for (const cmd of [
    "officecli view a.docx; Remove-Item C:\\x",
    "officecli view a.docx && calc",
    "officecli view a.docx || calc",
    "officecli view a.docx > out.txt",
    "officecli view a.docx &",
    "officecli view a.docx\ncalc",
    "officecli set a.docx / --prop text=$(calc)",
    'officecli set a.docx / --prop text="$(calc)"',
    "officecli set a.docx / --prop text=`calc`",
    'officecli get "a\\" ; calc ; echo \\""',
    "echo x | calc",
    "calc | officecli batch a.xlsx",
    "officecli view a.docx | Out-File x.txt",
    "echo hi",
    "officecli view a.docx |",
    "(officecli view a.docx)",
    "officecli view 'unterminated",
  ]) {
    assert.equal(isOfficecliOnlyShell(shell(cmd)), false, cmd);
  }
});

test("the SDK's own analysis can veto", () => {
  const ok = "officecli view a.docx outline";
  assert.equal(isOfficecliOnlyShell(shell(ok, { hasWriteFileRedirection: true })), false);
  assert.equal(isOfficecliOnlyShell(shell(ok, { possibleUrls: [{ url: "https://x" }] })), false);
  assert.equal(isOfficecliOnlyShell(shell(ok, { managedApprovalRequired: true })), false);
  assert.equal(isOfficecliOnlyShell(shell(ok, { requestSandboxBypass: true })), false);
});

test("works with the request shape the SDK sends for PowerShell (whole line as one identifier)", () => {
  const pipe = `echo '[{"command":"set","path":"/Sheet1/A1","props":{"value":"x"}}]' | officecli batch "C:\\ws\\p.xlsx" --json`;
  assert.ok(isOfficecliOnlyShell(shell(pipe, { commands: [{ identifier: pipe, readOnly: false }] })));
  const chained = 'officecli view "C:\\ws\\p.xlsx" outline; whoami';
  assert.equal(isOfficecliOnlyShell(shell(chained, { commands: [{ identifier: chained, readOnly: false }] })), false);
  const redirected = 'officecli view "C:\\ws\\p.xlsx" outline > "C:\\ws\\out.txt"';
  assert.equal(isOfficecliOnlyShell(shell(redirected, { commands: [{ identifier: redirected, readOnly: false }] })), false);
});

test("screenshots through the shell must use the HTML renderer to be auto-approved", () => {
  assert.equal(isOfficecliOnlyShell(shell("officecli view a.pptx screenshot -o o.png")), false);
  assert.equal(isOfficecliOnlyShell(shell("officecli view a.pptx stats --render native")), false);
  assert.ok(isOfficecliOnlyShell(shell("officecli view a.pptx screenshot -o o.png --render html")));
  assert.ok(isOfficecliOnlyShell(shell("officecli view a.pptx screenshot --render=html")));
});

test("the officecli tool still forces the HTML renderer", () => {
  assert.deepEqual(forceHtmlRenderer(["view", "a.pptx", "screenshot"]), ["view", "a.pptx", "screenshot", "--render", "html"]);
  assert.deepEqual(forceHtmlRenderer(["view", "a.pptx", "stats", "--render", "auto"]), ["view", "a.pptx", "stats", "--render", "html"]);
  assert.deepEqual(forceHtmlRenderer(["view", "a.pptx", "stats", "--render=native"]), ["view", "a.pptx", "stats", "--render=html"]);
  assert.deepEqual(forceHtmlRenderer(["view", "a.pptx", "outline"]), ["view", "a.pptx", "outline"]);
  assert.equal(usesNativeRenderer(["view", "a.pptx", "screenshot", "--render", "HTML"]), false);
});
