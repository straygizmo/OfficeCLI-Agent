# Composer "+" Menu, Attachments, Web Search, MCP/Skills Toggles — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "+" menu to the chat composer that attaches files (drag-drop, picker, paste), toggles Copilot web search, and toggles the MCP servers and skills the user's Copilot CLI already has, with toggle-switch styling throughout the settings panel.

**Architecture:** The React task pane sends attachments as base64 inside `chat.send`; the Node host saves them under the session workspace, lists their paths in the prompt and passes images as SDK blob attachments. The host enables Copilot config discovery so `~/.copilot` MCP servers and skills load, reports them as a `catalog`, and flips them live through the SDK session RPCs while persisting the choice in `~/.officecli-agent/config.json`. The C# relay stops JSON-parsing large page frames and gains an `openPath` bridge op.

**Tech Stack:** TypeScript (Node 20.19+, `node:test` via tsx), React 18 + Vite, `@github/copilot-sdk` 1.0.22, C# .NET Framework 4.8 VSTO (xunit via `dotnet test`), pnpm 9.

**Spec:** `docs/superpowers/specs/2026-10-21-composer-menu-attachments-mcp-skills-design.md`

## Global Constraints

- Attachment limits: at most **5 files per message, 10 MB (10 × 1024 × 1024 bytes) per file**; images are `image/png`, `image/jpeg`, `image/gif`, `image/webp`, `image/bmp`; everything else is a document passed by path only.
- Attachments are saved under `<workspaceDir>/attachments/`; names sanitized and de-duplicated as `report.docx`, `report (2).docx`, `report (3).docx`.
- The pane never writes anything under `~/.copilot` (honour `COPILOT_HOME` when set, else `~/.copilot`).
- Web search default **on**; off = `excludedTools: ["web_search", "web_fetch", "github-mcp-server-web_search"]` (fallbacks in Task 10).
- Session creation sets `enableConfigDiscovery: true` and `enableSkills: true`.
- MCP/skill toggles apply live (`session.rpc.mcp.enable/disable`, `session.rpc.skills.enable/disable`) **and** persist in `config.json` (`disabledMcpServers`, `disabledSkills`).
- Web search toggle: if the conversation has no user message, re-send `hello` immediately; otherwise show the "applies to the next new chat" status.
- All user-visible strings go through `packages/ui/src/i18n.ts` in both `ja` and `en`.
- Every checkbox in the settings panel becomes a `Toggle` (role="switch").
- The C# relay must not `JObject.Parse` frames other than bridge (`{"__bridge"`) and `hello` (`{"type":"hello"`) frames.
- Commit after every task; `pnpm typecheck` must pass before each commit that touches TypeScript.

## Review Focus

- A dropped file named like `../..\evil<>:"|?*.docx` or with only reserved characters must land inside `attachments/` with a safe name (test in Task 2).
- Two attachments with the same name in one message must both be saved, the second as `name (2).ext` (test in Task 2).
- A declared `size` that disagrees with the decoded bytes must be judged by the decoded bytes, and a decoded file over 10 MB refused (test in Task 2).
- When `mcp.discover` or `skills.list` throws (older runtime, MCP disabled by policy), `ready` must still arrive with an empty catalog instead of the session failing (test in Task 5).
- A page frame that is not JSON at all must be classified `Other` and forwarded unchanged, never crash the relay (test in Task 9).

---

### Task 1: Protocol additions

**Files:**
- Modify: `packages/protocol/src/index.ts`

**Interfaces:**
- Produces (used by every later task):

```ts
export const ATTACHMENT_MAX_FILES = 5;
export const ATTACHMENT_MAX_BYTES = 10 * 1024 * 1024;
export const IMAGE_MIME_TYPES: readonly string[] = ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"];
export const WEB_TOOL_NAMES: readonly string[] = ["web_search", "web_fetch", "github-mcp-server-web_search"];

export interface AttachmentInput { name: string; mimeType: string; size: number; dataBase64: string }

export interface McpServerSummary { name: string; type?: "stdio" | "http" | "sse"; source: string; enabled: boolean; cliDisabled: boolean }
export interface SkillSummary { name: string; description?: string; source: string; enabled: boolean; cliDisabled: boolean }
export interface Catalog { mcpServers: McpServerSummary[]; skills: SkillSummary[]; mcpConfigPath: string; skillsDir: string }
```

- `SessionSettings` gains `webSearch?: boolean` (doc comment: "Allow Copilot's web search / URL fetch tools. Default true. Applies to new sessions.").
- `ClientMessage`: `chat.send` becomes `{ type: "chat.send"; text: string; attachments?: AttachmentInput[] }`; add `{ type: "mcp.setEnabled"; name: string; enabled: boolean }` and `{ type: "skills.setEnabled"; name: string; enabled: boolean }`.
- `ServerMessage`: `ready` gains `catalog: Catalog`; add `{ type: "catalog"; catalog: Catalog }`.
- `AgentEvent` gains `{ kind: "server-tool"; name: string; outputIndex: number; status: "in_progress" | "searching" | "completed" }`.

- [ ] **Step 1: Add the types and constants above to `packages/protocol/src/index.ts`** (constants next to `COPILOT_PROVIDER`; summaries under the "Host -> task pane" banner).

- [ ] **Step 2: Build the protocol package and typecheck the workspace**

Run: `pnpm --filter @officecli-agent/protocol build && pnpm typecheck`
Expected: protocol builds; the host typecheck fails only in `server.ts` (`ready` missing `catalog`), which Task 5 fixes. If anything else fails, fix it here.

- [ ] **Step 3: Commit**

```bash
git add packages/protocol/src/index.ts
git commit -m "feat(protocol): attachments, web search setting, catalog and toggle messages"
```

---

### Task 2: Host attachment storage module

**Files:**
- Create: `packages/host/src/attachments.ts`
- Create: `packages/host/test/attachments.test.ts`
- Modify: `packages/host/package.json` (add `"test": "node --import tsx --test test/*.test.ts"`)

**Interfaces:**
- Consumes: `AttachmentInput`, `ATTACHMENT_MAX_FILES`, `ATTACHMENT_MAX_BYTES`, `IMAGE_MIME_TYPES` from Task 1.
- Produces:

```ts
export interface SavedAttachment { name: string; path: string; mimeType: string; size: number; isImage: boolean; dataBase64: string }
export function sanitizeFileName(name: string): string;              // strips directories, `<>:"/\|?*`, control chars; trims dots/spaces; "" → "attachment"
export function uniquePath(dir: string, fileName: string, exists: (p: string) => boolean): string; // "x.docx" → "x (2).docx" → "x (3).docx"
export function isImageMime(mimeType: string): boolean;
export function validateAttachments(list: AttachmentInput[]): string | undefined; // error text, or undefined when ok
export async function saveAttachments(workspaceDir: string, list: AttachmentInput[]): Promise<SavedAttachment[]>; // throws Error(validateAttachments(...)) when invalid
export function attachmentFooter(saved: SavedAttachment[]): string;  // "" for []; otherwise "\n\nAttached files (saved on disk):\n- <path> (<mimeType>)" per file
```

- [ ] **Step 1: Add the `test` script to `packages/host/package.json` and write the failing tests in `packages/host/test/attachments.test.ts`** (`node:test` + `node:assert/strict`; use `fs.mkdtempSync(path.join(os.tmpdir(), "oca-att-"))` for a workspace):

```ts
test("sanitizeFileName strips directories and reserved characters", () => {
  assert.equal(sanitizeFileName('../..\\evil<>:"|?*.docx'), "evil.docx");
  assert.equal(sanitizeFileName("<>:\"|?*"), "attachment");
  assert.equal(sanitizeFileName("  report .docx  "), "report .docx");
});
test("uniquePath numbers duplicates", () => {
  const taken = new Set(["/w/report.docx", "/w/report (2).docx"]);
  assert.equal(uniquePath("/w", "report.docx", (p) => taken.has(p)), path.join("/w", "report (3).docx"));
  assert.equal(uniquePath("/w", "new.png", () => false), path.join("/w", "new.png"));
});
test("validateAttachments enforces count and size", () => {
  const one = { name: "a.png", mimeType: "image/png", size: 1, dataBase64: "AA==" };
  assert.equal(validateAttachments([one]), undefined);
  assert.match(validateAttachments(Array(6).fill(one))!, /5/);
  assert.match(validateAttachments([{ ...one, size: ATTACHMENT_MAX_BYTES + 1 }])!, /10 MB/);
});
test("saveAttachments writes under attachments/, judges size by decoded bytes, keeps duplicates apart", async () => {
  const dir = mkdtemp();
  const data = Buffer.from("hello").toString("base64");
  const saved = await saveAttachments(dir, [
    { name: "a.txt", mimeType: "text/plain", size: 999, dataBase64: data },
    { name: "a.txt", mimeType: "text/plain", size: 5, dataBase64: data },
    { name: "pic.png", mimeType: "image/png", size: 5, dataBase64: data },
  ]);
  assert.deepEqual(saved.map((s) => path.basename(s.path)), ["a.txt", "a (2).txt", "pic.png"]);
  assert.equal(saved[0].size, 5);
  assert.equal(saved[2].isImage, true);
  assert.ok(saved[0].path.startsWith(path.join(dir, "attachments")));
  assert.equal(fs.readFileSync(saved[1].path, "utf8"), "hello");
});
test("saveAttachments refuses a decoded file over the limit", async () => {
  const big = Buffer.alloc(ATTACHMENT_MAX_BYTES + 1).toString("base64");
  await assert.rejects(saveAttachments(mkdtemp(), [{ name: "big.bin", mimeType: "application/octet-stream", size: 1, dataBase64: big }]), /10 MB/);
});
test("attachmentFooter lists absolute paths and mime types", () => {
  assert.equal(attachmentFooter([]), "");
  const text = attachmentFooter([{ name: "r.docx", path: "C:\\w\\attachments\\r.docx", mimeType: "application/x", size: 1, isImage: false, dataBase64: "" }]);
  assert.match(text, /^\n\nAttached files \(saved on disk\):\n- C:\\w\\attachments\\r\.docx \(application\/x\)$/);
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `pnpm --filter @officecli-agent/host test`
Expected: FAIL, cannot find module `../src/attachments.js`.

- [ ] **Step 3: Implement `packages/host/src/attachments.ts` with the signatures above.** `saveAttachments` creates `attachments/` with `fs.mkdir({ recursive: true })`, decodes with `Buffer.from(dataBase64, "base64")`, checks `buffer.length <= ATTACHMENT_MAX_BYTES`, and returns `size = buffer.length`.

- [ ] **Step 4: Run the tests to verify they pass**

Run: `pnpm --filter @officecli-agent/host test`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add packages/host/package.json packages/host/src/attachments.ts packages/host/test/attachments.test.ts
git commit -m "feat(host): save chat attachments under the session workspace"
```

---

### Task 3: Host sends attachments to the agent

**Files:**
- Modify: `packages/host/src/agent-session.ts` (`send`, event forwarding)
- Modify: `packages/host/src/system-prompt.ts`
- Modify: `packages/host/src/server.ts` (`chat.send` case, `WebSocketServer` options)

**Interfaces:**
- Consumes: `saveAttachments`, `attachmentFooter` (Task 2); `AttachmentInput`, `ATTACHMENT_MAX_FILES`, `ATTACHMENT_MAX_BYTES` (Task 1).
- Produces: `AgentSession.send(text: string, attachments?: AttachmentInput[]): Promise<void>`.

- [ ] **Step 1: Change `AgentSession.send` to `send(text, attachments = [])`.** When `attachments.length > 0`: `const saved = await saveAttachments(this.workspaceDir, attachments)`; prompt = `text + attachmentFooter(saved)`; SDK attachments = `saved.filter(s => s.isImage).map(s => ({ type: "blob", data: s.dataBase64, mimeType: s.mimeType, displayName: s.name }))`; call `this.session.send({ prompt, attachments: blobs.length ? blobs : undefined })`.

- [ ] **Step 2: Forward hosted web search progress.** In `forward()`, add `case "assistant.server_tool_progress": this.emit({ kind: "server-tool", name: event.data.kind, outputIndex: event.data.outputIndex, status: event.data.status as "in_progress" | "searching" | "completed" })`.

- [ ] **Step 3: Add the attachments paragraph to `buildSystemPrompt`** after the `office_insert_image` line:

```
- Attached files: files the user attaches to a message are saved under ${workspaceDir}\attachments\ and listed at the end of that message with absolute paths. Read Office files with officecli (never guess their contents). Images are shown to you directly when the model supports vision; to place one in the document use office_insert_image with its path.
```

- [ ] **Step 4: In `server.ts`** change the `chat.send` case to `await current.send(message.text, message.attachments)` and construct the server as `new WebSocketServer({ server: httpServer, path: WS_PATH, maxPayload: Math.ceil(ATTACHMENT_MAX_FILES * ATTACHMENT_MAX_BYTES * 4 / 3) + 1024 * 1024 })`.

- [ ] **Step 5: Typecheck and run tests**

Run: `pnpm typecheck && pnpm --filter @officecli-agent/host test`
Expected: tests pass; the only typecheck error left is the `ready.catalog` one from Task 1.

- [ ] **Step 6: Commit**

```bash
git add packages/host/src/agent-session.ts packages/host/src/system-prompt.ts packages/host/src/server.ts
git commit -m "feat(host): pass attachments to the agent and forward web search progress"
```

---

### Task 4: Config: disabled lists and Copilot CLI paths

**Files:**
- Modify: `packages/host/src/config.ts`
- Create: `packages/host/test/config.test.ts`

**Interfaces:**
- Produces:

```ts
// HostConfig gains (defaults [] / []):
disabledMcpServers: string[];   // Copilot CLI MCP servers switched off from the pane
disabledSkills: string[];       // skills switched off from the pane
export function withName(list: string[], name: string, present: boolean): string[]; // pure: add/remove, no duplicates, order kept
export function setMcpServerEnabled(config: HostConfig, name: string, enabled: boolean): void; // updates list, saveConfig
export function setSkillEnabled(config: HostConfig, name: string, enabled: boolean): void;
export function copilotHome(): string;            // process.env.COPILOT_HOME || ~/.copilot
export function copilotMcpConfigPath(): string;   // <home>/mcp-config.json
export function copilotSkillsDir(): string;       // <home>/skills
```

- [ ] **Step 1: Write failing tests in `packages/host/test/config.test.ts`**

```ts
test("withName adds once and removes", () => {
  assert.deepEqual(withName([], "a", true), ["a"]);
  assert.deepEqual(withName(["a"], "a", true), ["a"]);
  assert.deepEqual(withName(["a", "b"], "a", false), ["b"]);
  assert.deepEqual(withName(["b"], "a", false), ["b"]);
});
test("copilot paths honour COPILOT_HOME", () => {
  process.env.COPILOT_HOME = "C:\\ch";
  assert.equal(copilotMcpConfigPath(), path.join("C:\\ch", "mcp-config.json"));
  assert.equal(copilotSkillsDir(), path.join("C:\\ch", "skills"));
  delete process.env.COPILOT_HOME;
  assert.equal(copilotHome(), path.join(os.homedir(), ".copilot"));
});
test("loadConfig defaults the disabled lists to empty arrays", () => {
  const cfg = loadConfig({ disabledMcpServers: undefined as unknown as string[] });
  assert.deepEqual(cfg.disabledMcpServers, []);
  assert.deepEqual(cfg.disabledSkills, []);
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pnpm --filter @officecli-agent/host test`
Expected: FAIL on missing exports.

- [ ] **Step 3: Implement in `config.ts`.** Add the two fields to `HostConfig` and `defaults`; in `loadConfig` coerce both to arrays (`Array.isArray(x) ? x : []`) like `providers`; `setMcpServerEnabled` = `config.disabledMcpServers = withName(config.disabledMcpServers, name, !enabled); saveConfig(config)`; same for skills.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pnpm --filter @officecli-agent/host test`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add packages/host/src/config.ts packages/host/test/config.test.ts
git commit -m "feat(host): persist disabled MCP servers and skills; locate the Copilot CLI config"
```

---

### Task 5: Host catalog, discovery, live toggles, web search exclusion

**Files:**
- Create: `packages/host/src/catalog.ts`
- Create: `packages/host/test/catalog.test.ts`
- Modify: `packages/host/src/agent-session.ts` (`start`, new methods)
- Modify: `packages/host/src/server.ts` (`ready`, new cases, `settings.update` note)

**Interfaces:**
- Consumes: Task 1 types; Task 4 helpers; SDK `client.rpc.mcp.discover({})` → `{ servers: { name, type?, source, enabled }[] }`; `session.rpc.skills.list()` → `{ skills: { name, description, source, enabled }[] }`; `session.rpc.mcp.enable({ serverName })` / `disable({ serverName })`; `session.rpc.skills.enable({ name })` / `disable({ name })`.
- Produces:

```ts
// catalog.ts
export function mergeMcpServers(discovered: { name: string; type?: string; source: string; enabled: boolean }[], disabled: string[]): McpServerSummary[];
export function mergeSkills(skills: { name: string; description?: string; source: string; enabled: boolean }[], disabled: string[]): SkillSummary[];
export async function buildCatalog(discover: () => Promise<{ servers: ... }>, listSkills: () => Promise<{ skills: ... }>, config: HostConfig): Promise<Catalog>; // each source falls back to [] on throw (console.warn)
// agent-session.ts
get link(): TaskpaneLink;
get catalog(): Catalog;                         // set during start()
async refreshCatalog(): Promise<Catalog>;
async setMcpEnabled(name: string, enabled: boolean): Promise<void>;
async setSkillEnabled(name: string, enabled: boolean): Promise<void>;
```

Rules: MCP `enabled = d.enabled && !disabled.includes(d.name)`, `cliDisabled = !d.enabled`, `type` kept only when it is `"stdio" | "http" | "sse"`. Skills: `cliDisabled = !s.enabled && !disabled.includes(s.name)`, `enabled = s.enabled`. Both lists sorted by name.

- [ ] **Step 1: Write failing tests in `packages/host/test/catalog.test.ts`**

```ts
test("mergeMcpServers applies the pane's disabled list on top of the CLI's", () => {
  const out = mergeMcpServers([{ name: "gh", type: "http", source: "user", enabled: true }, { name: "fs", type: "stdio", source: "user", enabled: false }], ["gh"]);
  assert.deepEqual(out, [
    { name: "fs", type: "stdio", source: "user", enabled: false, cliDisabled: true },
    { name: "gh", type: "http", source: "user", enabled: false, cliDisabled: false },
  ]);
});
test("mergeSkills tells CLI-disabled from pane-disabled", () => {
  const out = mergeSkills([{ name: "a", description: "A", source: "user", enabled: false }, { name: "b", description: "B", source: "builtin", enabled: false }], ["a"]);
  assert.equal(out[0].cliDisabled, false);
  assert.equal(out[1].cliDisabled, true);
});
test("buildCatalog survives a failing source", async () => {
  const cfg = { disabledMcpServers: [], disabledSkills: [] } as unknown as HostConfig;
  const cat = await buildCatalog(() => Promise.reject(new Error("no mcp")), async () => ({ skills: [{ name: "s", description: "", source: "user", enabled: true }] }), cfg);
  assert.deepEqual(cat.mcpServers, []);
  assert.equal(cat.skills[0].name, "s");
  assert.equal(cat.mcpConfigPath, copilotMcpConfigPath());
  assert.equal(cat.skillsDir, copilotSkillsDir());
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `pnpm --filter @officecli-agent/host test`
Expected: FAIL on missing module.

- [ ] **Step 3: Implement `catalog.ts`** with the signatures and rules above.

- [ ] **Step 4: Run tests to verify they pass**

Run: `pnpm --filter @officecli-agent/host test`
Expected: all pass.

- [ ] **Step 5: Wire `AgentSession`.** In `start()` add to `createSession`: `enableConfigDiscovery: true`, `enableSkills: true`, `disabledMcpServers: config.disabledMcpServers`, `disabledSkills: config.disabledSkills`, and `excludedTools: this.settings.webSearch === false ? [...WEB_TOOL_NAMES] : undefined`. After the session exists: `this.catalogValue = await buildCatalog(() => client.rpc.mcp.discover({}), () => session.rpc.skills.list(), config)`. `refreshCatalog()` repeats that and returns it. `setMcpEnabled` calls `session.rpc.mcp[enabled ? "enable" : "disable"]({ serverName: name })`; `setSkillEnabled` calls `session.rpc.skills[enabled ? "enable" : "disable"]({ name })`; both swallow and log errors with `console.warn` (the config change still persists).

- [ ] **Step 6: Wire `server.ts`.** In `ready` add `catalog: next.catalog` and `webSearch: message.settings?.webSearch ?? true` under `settings`. Add cases:

```ts
case "mcp.setEnabled":
  setMcpServerEnabled(config, message.name, message.enabled);
  await current?.setMcpEnabled(message.name, message.enabled);
  await broadcastCatalog();
  return current;
case "skills.setEnabled": /* same with setSkillEnabled / setSkillEnabled */
```

`broadcastCatalog()` (closure inside `startHost`) does `for (const s of sessions) s.link.send({ type: "catalog", catalog: await s.refreshCatalog() })`. In `settings.update`, add `message.settings.webSearch !== undefined` to the condition and change the text to `"Provider, model, open-file access and web search changes apply to the next new chat."`.

- [ ] **Step 7: Typecheck and run tests**

Run: `pnpm typecheck && pnpm --filter @officecli-agent/host test`
Expected: typecheck clean across all packages; all host tests pass.

- [ ] **Step 8: Commit**

```bash
git add packages/host/src/catalog.ts packages/host/test/catalog.test.ts packages/host/src/agent-session.ts packages/host/src/server.ts
git commit -m "feat(host): load Copilot CLI MCP servers and skills, report a catalog, live toggles, web search exclusion"
```

---

### Task 6: Toggle component and settings switches

**Files:**
- Create: `packages/ui/src/Toggle.tsx`
- Modify: `packages/ui/src/App.tsx` (the two `label.row` checkboxes in the settings panel)
- Modify: `packages/ui/src/styles.css`

**Interfaces:**
- Produces:

```tsx
export function Toggle(props: { checked: boolean; onChange(next: boolean): void; label?: React.ReactNode; disabled?: boolean; title?: string; className?: string }): JSX.Element;
```

Markup: `<label class="toggle"><input type="checkbox" role="switch" aria-checked … /><span class="toggle-track"><span class="toggle-knob"/></span>{label && <span class="toggle-label">{label}</span>}</label>`. The input is visually hidden but focusable; `:focus-visible` on the input outlines the track.

- [ ] **Step 1: Create `Toggle.tsx`** as specified.

- [ ] **Step 2: Add styles** to `styles.css`: track 32×18 px, radius 999px, background `var(--border)` off / `var(--accent)` on; knob 14 px white circle translating 14 px when on; `transition: 0.15s`; disabled at `opacity: 0.5`; works in dark mode (uses the existing tokens only).

- [ ] **Step 3: Replace the two settings checkboxes** (`autoApproveReadOnly`, `workAcrossOpenFiles`) with `<Toggle checked=… onChange=… label={t.…} />`; keep the `hint` paragraph below `workAcrossOpenFiles` and adjust `.settings .hint` margin so it still aligns under the label.

- [ ] **Step 4: Build the UI and check in the browser**

Run: `pnpm build` then `pnpm start` and open `http://localhost:4320/taskpane.html?hostApp=Word` → settings.
Expected: both rows show switches; clicking flips them and the status line still reports the "next new chat" note for open-file access; keyboard Tab + Space toggles.

- [ ] **Step 5: Commit**

```bash
git add packages/ui/src/Toggle.tsx packages/ui/src/App.tsx packages/ui/src/styles.css
git commit -m "feat(ui): toggle switches in the settings panel"
```

---

### Task 7: "+" menu with web search, MCP and skills submenus

**Files:**
- Create: `packages/ui/src/PlusMenu.tsx`
- Modify: `packages/ui/src/App.tsx` (catalog state, composer row, handlers)
- Modify: `packages/ui/src/host-client.ts` (`openPath`)
- Modify: `packages/ui/src/i18n.ts`
- Modify: `packages/ui/src/styles.css`

**Interfaces:**
- Consumes: `Catalog`, `SessionSettings.webSearch`, `mcp.setEnabled`, `skills.setEnabled`, `catalog` message (Task 1/5); `Toggle` (Task 6).
- Produces:

```tsx
// PlusMenu.tsx
export function PlusMenu(props: {
  disabled: boolean;                // state !== "ready"
  webSearch: boolean;
  onWebSearch(next: boolean): void;
  catalog: Catalog | null;
  onMcpToggle(name: string, enabled: boolean): void;
  onSkillToggle(name: string, enabled: boolean): void;
  onAddFiles(): void;
  onOpenPath(path: string): void;
}): JSX.Element;
// host-client.ts
HostClient.openPath(path: string): boolean;   // embedded: posts {__bridge:"openPath", path} and returns true; browser: returns false
```

Menu rows in order: `t.addFiles` (paperclip icon), `t.searchWeb` (globe icon, check mark at right when on; the row itself toggles), `t.mcpServers` ›, `t.skills` ›. Submenu view: back row (‹ title), then one row per item with the name (and `description` muted for skills, `source` badge for both) and a `Toggle` (disabled when `cliDisabled`, title `t.cliDisabledTip`); empty list → `t.noMcpServers` / `t.noSkills`; last row `t.openCopilotSettings` which calls `onOpenPath(catalog.mcpConfigPath | catalog.skillsDir)`. Menu closes on Escape, click outside, and after `onAddFiles`. Popover anchored above the "+" button (`position: absolute; bottom: 100%; left: 0`), max-height 60vh with scroll.

- [ ] **Step 1: Add i18n keys** to both `ja` and `en`: `addFiles` ("ファイルや写真を追加" / "Add files or photos"), `searchWeb` ("Web を検索" / "Search the web"), `mcpServers` ("MCP サーバー" / "MCP servers"), `skills` ("スキル" / "Skills"), `noMcpServers` ("Copilot CLI に MCP サーバーが設定されていません" / "No MCP servers configured in Copilot CLI"), `noSkills` ("Copilot CLI にスキルがありません" / "No skills found in Copilot CLI"), `openCopilotSettings` ("Copilot CLI の設定を開く…" / "Open Copilot CLI settings…"), `cliDisabledTip` ("Copilot CLI 側で無効になっています" / "Disabled in Copilot CLI itself"), `plusMenu` ("メニュー" / "Menu"), `back` exists already, `webSearchNextChat` ("Web 検索の変更は次の新しいチャットから有効になります" / "The web search change applies to the next new chat"), `openPathHint: (p: string) => …` ("ここを開いてください: " + p / "Open this location: " + p), `copilotMcpConfig` ("Copilot CLI MCP 設定" / "Copilot CLI MCP config"), `copilotSkillsDir` ("Copilot CLI スキル" / "Copilot CLI skills").

- [ ] **Step 2: Add `openPath` to `HostClient`** (returns false when not embedded).

- [ ] **Step 3: Create `PlusMenu.tsx`** as specified, using `useEffect` for the document `mousedown` / `keydown` listeners while open.

- [ ] **Step 4: Wire `App.tsx`.** Add `const [catalog, setCatalog] = useState<Catalog | null>(null)`; set it from `ready` (`message.catalog`) and from a new `case "catalog"`. Handlers:
  - `onWebSearch(next)`: `updateSession({ webSearch: next })`; then if `!items.some(i => i.role === "user")` → `clientRef.current?.sendHello(); setActiveSession(…)`, else `setStatus(t.webSearchNextChat)`.
  - `onMcpToggle`/`onSkillToggle`: `clientRef.current?.send({ type: "mcp.setEnabled" | "skills.setEnabled", name, enabled })` and optimistically update `catalog`.
  - `onOpenPath(p)`: `if (!clientRef.current?.openPath(p)) setStatus(t.openPathHint(p))`.
  - `onAddFiles`: `fileInputRef.current?.click()` (the input itself is added in Task 8; add the ref and a hidden `<input type="file" multiple ref=…>` now with an `onChange` that is filled in Task 8).
  Render `<PlusMenu …/>` as the first child of a new `div.composer-row` that wraps the existing `div.input-wrap` (the input-wrap keeps the send/stop button).
  In the settings `info` block add two lines: `{t.copilotMcpConfig}: <code>{catalog.mcpConfigPath}</code>` and `{t.copilotSkillsDir}: <code>{catalog.skillsDir}</code>`.

- [ ] **Step 5: Styles.** `.composer-row { display: flex; gap: 6px; align-items: flex-end; }`, `.plus-wrap { position: relative; }`, `.plus-menu { position: absolute; bottom: 100%; left: 0; margin-bottom: 6px; min-width: 240px; max-width: 320px; max-height: 60vh; overflow-y: auto; background: var(--bg); border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 6px 24px rgba(0,0,0,.18); padding: 4px; z-index: 10; }`, `.menu-row` (flex, 8px gap, 8px 10px padding, radius 6px, hover `var(--assistant)`, full-width button reset), `.menu-row .grow { flex: 1 }`, `.menu-check { color: var(--accent) }`, `.menu-sub { opacity: .6 }`, `.source-badge` (11px, muted, border, radius 999px, 0 6px).

- [ ] **Step 6: Build and check in the browser**

Run: `pnpm build && pnpm start`, open the debug page.
Expected: "+" opens the menu; web search row shows a check and toggling it with an empty chat restarts the session (status "Starting Copilot session…"); MCP submenu lists the servers from `~/.copilot/mcp-config.json` (or the empty hint); skills submenu lists at least `officecli` (from `~/.copilot/skills`) plus the runtime's built-in skills; switching one off and reopening shows it off; "Open Copilot CLI settings…" shows the path in the status line (browser mode).

- [ ] **Step 7: Commit**

```bash
git add packages/ui/src/PlusMenu.tsx packages/ui/src/App.tsx packages/ui/src/host-client.ts packages/ui/src/i18n.ts packages/ui/src/styles.css
git commit -m "feat(ui): composer plus menu with web search, MCP server and skill toggles"
```

---

### Task 8: Attachments in the composer and chat

**Files:**
- Create: `packages/ui/src/attachments.ts`
- Create: `packages/ui/src/AttachmentChips.tsx`
- Modify: `packages/ui/src/App.tsx` (pending state, drop/paste/picker, send, `ChatItem`, `ChatRow`, server-tool rows)
- Modify: `packages/ui/src/i18n.ts`
- Modify: `packages/ui/src/styles.css`

**Interfaces:**
- Consumes: `AttachmentInput`, `ATTACHMENT_MAX_FILES`, `ATTACHMENT_MAX_BYTES`, `IMAGE_MIME_TYPES` (Task 1); `server-tool` event (Task 3).
- Produces:

```ts
// attachments.ts
export interface PendingAttachment { id: string; file: File; previewUrl?: string }   // previewUrl only for images (URL.createObjectURL)
export interface ShownAttachment { name: string; mimeType: string; size: number; thumbnail?: string } // stored in ChatItem / history
export function isImageFile(file: File): boolean;                 // by IMAGE_MIME_TYPES, falling back to extension when file.type is ""
export function acceptFiles(current: PendingAttachment[], incoming: File[]): { next: PendingAttachment[]; error?: string }; // enforces 5 / 10 MB, error text via t
export function readAsBase64(file: File): Promise<string>;        // FileReader → strip the data: prefix
export function makeThumbnail(file: File, maxPx?: number): Promise<string | undefined>; // canvas, default 64, undefined for non-images / failures
export function formatSize(bytes: number): string;                // "12 KB", "3.4 MB"
// AttachmentChips.tsx
export function AttachmentChips(props: { items: Array<{ id?: string; name: string; size: number; previewUrl?: string; thumbnail?: string }>; onRemove?(id: string): void }): JSX.Element;
```

`ChatItem` user variant becomes `{ id; role: "user"; text; attachments?: ShownAttachment[] }`.

- [ ] **Step 1: Add i18n keys**: `tooManyFiles: (n) =>` ("添付できるのは ${n} 件までです" / "You can attach up to ${n} files"), `fileTooLarge: (name, mb) =>` ("${name} は ${mb} MB を超えています" / "${name} is larger than ${mb} MB"), `removeAttachment` ("添付を外す" / "Remove attachment"), `dropHere` ("ここにドロップして添付" / "Drop to attach"), `searchingWeb` ("Web を検索中…" / "Searching the web…").

- [ ] **Step 2: Create `attachments.ts` and `AttachmentChips.tsx`** as specified. Chips: image chips show the preview/thumbnail at 28 px; others show a file icon; name is ellipsized at 160 px; size muted; × button when `onRemove`.

- [ ] **Step 3: Wire the composer in `App.tsx`.** State `pending: PendingAttachment[]`, `dragOver: boolean`. On `footer.composer`: `onDragOver` (preventDefault, set dragOver when `e.dataTransfer.types.includes("Files")`), `onDragLeave`, `onDrop` (preventDefault; `acceptFiles(pending, Array.from(e.dataTransfer.files))`; set error into status). On the textarea `onPaste`: collect `clipboardData.items` of kind `file` with image type → `acceptFiles`. Hidden `<input type="file" multiple>` `onChange` → `acceptFiles`, then clear `e.target.value`. Render `<AttachmentChips items={pending} onRemove=…/>` above the textarea when non-empty; add class `drop-active` on the composer while dragging. Send button enabled when `input.trim() || pending.length`.

- [ ] **Step 4: Change `send()`** to async: build `attachments: AttachmentInput[]` with `readAsBase64` for each pending file, `shown: ShownAttachment[]` with `makeThumbnail`; append the user item with `attachments: shown`; send `{ type: "chat.send", text, attachments }`; revoke object URLs; clear `pending`. Allow an empty `text` when attachments exist.

- [ ] **Step 5: Render attachments in `ChatRow`** for `role: "user"`: `<AttachmentChips items={item.attachments}/>` under the text. History already serialises `items`, so thumbnails persist with the 64 px cap.

- [ ] **Step 6: Server-tool rows.** In `handleEvent` add `case "server-tool"`: id `server-tool-${event.outputIndex}`; on `in_progress`/`searching` upsert a tool item `{ role: "tool", toolName: t.searchingWeb, status: "running" }`; on `completed` set status `ok`.

- [ ] **Step 7: Styles.** `.chips { display: flex; flex-wrap: wrap; gap: 4px; margin-bottom: 4px; }`, `.chip { display: inline-flex; align-items: center; gap: 6px; border: 1px solid var(--border); border-radius: 8px; padding: 3px 6px; font-size: 11px; max-width: 100%; }`, `.chip img { width: 28px; height: 28px; object-fit: cover; border-radius: 4px; }`, `.chip .name { max-width: 160px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }`, `.composer.drop-active { outline: 2px dashed var(--accent); outline-offset: -4px; }`, `.msg.user .chips { margin-top: 6px; }`.

- [ ] **Step 8: Build and check in the browser**

Run: `pnpm build && pnpm start`, open the debug page.
Expected: dropping a png and a docx shows two chips; sending "What is in these files?" shows the chips in the user bubble, the assistant describes the image, and a tool row shows officecli reading the docx (the `attachments/` folder under the workspace contains both files). A sixth file or an 11 MB file is refused with the status text. Pasting a screenshot adds an image chip.

- [ ] **Step 9: Commit**

```bash
git add packages/ui/src/attachments.ts packages/ui/src/AttachmentChips.tsx packages/ui/src/App.tsx packages/ui/src/i18n.ts packages/ui/src/styles.css
git commit -m "feat(ui): attach files by drag-and-drop, picker and paste"
```

---

### Task 9: VSTO relay: no-parse forwarding and the openPath bridge

**Files:**
- Modify: `vsto/OfficeCliAgent.Core/HostRelay.cs` (`OnPageMessageAsync`, new static helper)
- Modify: `vsto/OfficeCliAgent.Core/AgentTaskPane.cs` (`AllowExternalDrop`)
- Create: `vsto/OfficeCliAgent.Core.Tests/HostRelayTests.cs`

**Interfaces:**
- Produces:

```csharp
public enum PageFrameKind { Bridge, Hello, Other }
public static PageFrameKind HostRelay.ClassifyPageFrame(string text); // prefix test only: "{\"__bridge\"" → Bridge, "{\"type\":\"hello\"" → Hello, else Other
public static string HostRelay.ResolveOpenTarget(string path, Func<string, bool> fileExists, Func<string, bool> dirExists); // path if it exists; else its parent directory if that exists; else null
```

- [ ] **Step 1: Write failing xunit tests in `HostRelayTests.cs`**

```csharp
[Theory]
[InlineData("{\"__bridge\":\"ready\"}", PageFrameKind.Bridge)]
[InlineData("{\"__bridge\":\"openPath\",\"path\":\"C:\\\\x\"}", PageFrameKind.Bridge)]
[InlineData("{\"type\":\"hello\",\"protocolVersion\":1}", PageFrameKind.Hello)]
[InlineData("{\"type\":\"chat.send\",\"text\":\"hi\",\"attachments\":[{\"dataBase64\":\"AAAA\"}]}", PageFrameKind.Other)]
[InlineData("not json at all", PageFrameKind.Other)]
[InlineData("", PageFrameKind.Other)]
public void ClassifyPageFrameLooksOnlyAtThePrefix(string text, PageFrameKind expected) => Assert.Equal(expected, HostRelay.ClassifyPageFrame(text));

[Fact]
public void ResolveOpenTargetFallsBackToTheParentFolder()
{
    Assert.Equal(@"C:\c\mcp-config.json", HostRelay.ResolveOpenTarget(@"C:\c\mcp-config.json", p => p == @"C:\c\mcp-config.json", _ => false));
    Assert.Equal(@"C:\c", HostRelay.ResolveOpenTarget(@"C:\c\mcp-config.json", _ => false, d => d == @"C:\c"));
    Assert.Null(HostRelay.ResolveOpenTarget(@"C:\nope\x.json", _ => false, _ => false));
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: build error, `PageFrameKind` / `ClassifyPageFrame` missing.

- [ ] **Step 3: Implement.** In `OnPageMessageAsync`: `switch (ClassifyPageFrame(text))` — `Other` → `await SendAsync(text)` without parsing; `Hello` → existing stamping; `Bridge` → parse, `ready` as today, `openPath` → `var target = ResolveOpenTarget(path, File.Exists, Directory.Exists); if (target != null) Process.Start(new ProcessStartInfo(target) { UseShellExecute = true })` inside try/catch with `Log.Write`. Update the class doc comment ("Frames flow through unchanged… large chat frames are never parsed here"). In `AgentTaskPane.InitializeAsync` set `_web.AllowExternalDrop = true` right after `EnsureCoreWebView2Async` (if the installed WebView2 package lacks the property, note it and rely on the default).

- [ ] **Step 4: Run the tests to verify they pass**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add vsto/OfficeCliAgent.Core/HostRelay.cs vsto/OfficeCliAgent.Core/AgentTaskPane.cs vsto/OfficeCliAgent.Core.Tests/HostRelayTests.cs
git commit -m "feat(vsto): forward large page frames unparsed; openPath bridge; allow external drop"
```

---

### Task 10: Real-add-in verification, web search fallback, docs

**Files:**
- Modify (only if the fallback is needed): `packages/host/src/agent-session.ts`, `packages/host/src/system-prompt.ts`, `packages/ui/src/i18n.ts`
- Modify: `README.md`, `README-ja.md`

**Interfaces:** none new.

- [ ] **Step 1: Build everything and load it into Word** following the memory note "Office loads the repo Debug build": `pnpm build`; close Word; MSBuild the three `.sln` files with `Configuration=Debug`; kill the node process listening on 127.0.0.1:4320; open Word and the Agent pane.

- [ ] **Step 2: Attachments on the real pane.** Drop a png and a docx onto the composer (chips appear; if the drop does nothing, check `AllowExternalDrop` and the WebView2 package version and fix in `AgentTaskPane.cs`). Send "What is in these files?". Expected: image described, officecli tool row for the docx. Then attach a 10 MB file and send: expected to arrive (host log shows the saved path). Pick a file through "Add files or photos": expected native dialog.

- [ ] **Step 3: Web search off.** In a fresh chat switch "Search the web" off (session restarts), ask "What is in the news today?". Expected: no `web_search` / `web_fetch` tool row and the answer says it cannot browse. If a search still runs:
  1. add `webSearch: false` to the session config (cast: `...(this.settings.webSearch === false ? ({ webSearch: false } as object) : {})`) and retest;
  2. if that also fails, append to the system prompt when off: `"- Web search is turned off by the user: do not call web_search or web_fetch and do not browse URLs."`, add " (best effort)" to `t.searchWeb` via a new key `searchWebBestEffort`, and record the outcome in the plan file.

- [ ] **Step 4: MCP and skills on the real pane.** Expected: the MCP submenu lists the servers from `~/.copilot/mcp-config.json`; the skills submenu lists `officecli`. Switch one MCP server off, ask the agent to use it: expected refusal/no such tool. Switch it on again in the same chat: expected usable. "Open Copilot CLI settings…" opens the file / folder in Windows.

- [ ] **Step 5: Document** in `README.md` and `README-ja.md` (a short "Composer menu" subsection): attachments (limits, where files go), web search toggle, MCP servers and skills shared with the Copilot CLI (how to add them there), and the two config keys `disabledMcpServers` / `disabledSkills`.

- [ ] **Step 6: Final checks and commit**

Run: `pnpm typecheck && pnpm --filter @officecli-agent/host test && dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: all pass.

```bash
git add README.md README-ja.md packages docs
git commit -m "docs: composer menu, attachments, web search and Copilot CLI MCP/skills"
```
