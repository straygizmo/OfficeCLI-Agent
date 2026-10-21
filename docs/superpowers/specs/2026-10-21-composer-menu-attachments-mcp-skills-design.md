# Composer "+" menu: attachments, web search, MCP servers, skills — design

Date: 2026-10-21

## Intent

The chat pane should offer the same "+" menu next to the message box that
the Claude desktop add-in has (reference screenshot: "Add files or photos",
"Search the web" with a check mark, "Connectors ›", "Plugins ›", "Skills ›"),
adapted to what GitHub Copilot CLI provides:

- attach images and Office files by drag-and-drop or a file picker, so the
  agent can look at them and work with them;
- switch Copilot's web search on and off;
- switch the MCP servers and skills that are configured for the user's
  GitHub Copilot CLI on and off for this agent.

MCP servers and skills are **shared with the Copilot CLI**: the pane shows
what the CLI has (`~/.copilot/mcp-config.json`, `~/.copilot/skills`, the
runtime's built-in skills) and only switches items on or off. Adding,
editing and removing them stays in the CLI's own configuration. This keeps
one source of truth and avoids a second, confusing configuration surface.

Today the pane has none of this: the agent only sees the live document,
MCP servers and skills configured for the Copilot CLI are not loaded at all
(`enableConfigDiscovery` is off), and web search is whatever the runtime
decides.

## Goals

- A "+" button at the left of the composer opens a menu with, in this order:
  "Add files or photos", "Search the web" (toggle), "MCP servers ›",
  "Skills ›". Each submenu lists the items with a toggle switch. The last
  row of each submenu, "Open Copilot CLI settings…", opens the CLI's
  configuration location (the MCP config file, or the skills folder) in
  Windows so the user can add or remove items there.
- Files can be attached by dropping them anywhere on the composer, by the
  menu entry (native file picker) or by pasting an image from the clipboard.
  Attached files show as chips (name, size, × to remove) above the text box
  and are sent together with the message. The user message in the chat shows
  the chips too (a small thumbnail for images).
- The agent can see attached images (vision) and can read attached Office
  files, PDFs and text files through its existing tools (officecli, file
  reads). Attached images can be inserted into the document through the
  existing `insertImage` tool because they exist on disk.
- "Search the web" off removes Copilot's hosted `web_search` and `web_fetch`
  tools from the session. On is the default (today's behaviour).
- MCP servers and skills configured for the Copilot CLI are loaded into the
  session (config discovery on) and every loaded item is visible in the
  lists. An item switched off is not started / not offered to the model.
  The on/off state is stored in `~/.officecli-agent/config.json`
  (shared by the Word, Excel and PowerPoint panes); the pane never writes
  to `~/.copilot`.
- Every checkbox in the settings panel (including the existing "auto-approve
  read-only" and "work across open files") is drawn as a toggle switch.
- All new text is in the i18n table (Japanese and English).

## Non-goals

- Adding, editing or deleting MCP servers or skills from the pane. The user
  does that with the Copilot CLI (`copilot mcp add …`, or editing
  `~/.copilot/mcp-config.json`, or putting a folder with `SKILL.md` under
  `~/.copilot/skills`).
- Installing MCP servers or skills from a marketplace / Agent Finder.
- Remote MCP servers that need an interactive OAuth login from inside the
  pane. If the CLI has already logged in and stored the token, the server
  works; otherwise it is reported as failed in the tool row.
- Attaching whole folders, or files over the size limit (see Constraints).
- Showing web search results or citations in a special way. A search shows
  up as a tool row like any other tool call.
- Plugins ("Open Plugins" directories) and custom agents as separate
  switchable lists. If discovery loads them, they simply apply.

## Anti-goals (what would make this a failure)

- The agent receives a `.docx` as pasted text garbage because the runtime
  tried to read a binary file into the prompt.
- Turning on config discovery silently changes what the agent can do without
  the user seeing it in the lists (every loaded MCP server and skill must be
  visible and switchable).
- A toggle in the "+" menu looks immediate but does nothing until the user
  guesses that a new chat is needed.
- The pane writes to anything under `~/.copilot`.
- A dropped file lands in a random place on disk or outside the session
  workspace.
- The feature works in the browser debugging page but not inside the
  WebView2 task pane (drag-and-drop, file picker and "open settings" must
  be verified in real Word).

## Constraints

- Attachments: at most 5 files per message, 10 MB per file. Larger files
  are refused with a status message. Images are png, jpg/jpeg, gif, webp,
  bmp. Anything else is treated as a document and goes to the agent by
  path.
- Attachments travel as base64 inside the `chat.send` frame. Inside Word
  the frame goes page → WebView2 `postMessage` → the C# relay → WebSocket,
  so a full message can be about 70 MB of JSON. The relay today parses
  every page frame with `JObject.Parse` to find bridge and `hello` frames;
  it must instead test the frame's first bytes (`{"__bridge"` /
  `{"type":"hello"`) and forward anything else untouched, so large frames
  are never parsed inside the Office process. The host's WebSocket server
  sets `maxPayload` explicitly to cover 5 × 10 MB × 4/3 plus slack. The
  builder verifies a 10 MB attachment in the real task pane. (My call: the
  lower limit plus no-parse relay beats a second native upload path.)
- Attachments are written under the session's workspace directory
  (`<workspaceDir>/attachments/`), which the agent is already allowed to
  read and write. A file name is sanitized and de-duplicated
  (`report.docx`, `report (2).docx`).
- MCP server and skill toggles apply **immediately to the running
  session**: the SDK offers `session.rpc.mcp.enable/disable` and
  `session.rpc.skills.enable/disable`, so the host calls those and also
  persists the choice in `config.json` for the next sessions. No "next
  chat" note for these. (My call; it avoids the restart dance.)
- The web search toggle has no live RPC; it is a session creation option.
  When the conversation on screen has no user message yet, the pane
  re-sends `hello` right away (same as "new chat"), so the toggle is live
  before the first message. When the conversation already has messages,
  the pane shows the existing "applies to the next new chat" note in the
  status line.
- `enableConfigDiscovery: true` and `enableSkills: true` must be set on
  session creation so the runtime loads the CLI's MCP servers and skills.
  The builder checks what else discovery loads from the workspace directory
  and the home directory (instructions, hooks, plugins) and notes it in the
  plan; the workspace directory is ours and empty, so this is expected to
  be only the user's own `~/.copilot` configuration.

## How it is built — decided

### Protocol (`packages/protocol`)

- `chat.send` gains `attachments?: { name: string; mimeType: string; size: number; dataBase64: string }[]`.
- `SessionSettings` gains `webSearch?: boolean` (default true). It lives
  with the other per-pane session settings (localStorage, sent in `hello`).
- New client messages: `mcp.setEnabled { name, enabled }` and
  `skills.setEnabled { name, enabled }`.
- New host → pane message `catalog { mcpServers: McpServerSummary[]; skills: SkillSummary[]; mcpConfigPath: string; skillsDir: string }`,
  also included in `ready`. Summaries never include secrets.
  - `McpServerSummary { name; type?: "stdio" | "http" | "sse"; source: string; enabled }`
  - `SkillSummary { name; description?; source: string; enabled }`
- `AgentEvent` gains `{ kind: "server-tool"; name: "web_search"; status: "in_progress" | "searching" | "completed" }`
  from the runtime's `assistant.server_tool_progress` event, so the pane can
  show a "Searching the web…" tool row.

### Host (`packages/host`)

- `config.json` gains `disabledMcpServers: string[]` and
  `disabledSkills: string[]`, with `setMcpServerEnabled` and
  `setSkillEnabled` helpers next to the provider helpers; both persist with
  `saveConfig`.
- `AgentSession.start()` passes `enableConfigDiscovery: true`,
  `enableSkills: true`, `disabledMcpServers`, `disabledSkills`, and, when
  `webSearch` is false, `excludedTools: ["web_search", "web_fetch"]`.
  Web search in the runtime can come from three places: the hosted
  (provider-side) `web_search`, the built-in `web_fetch` tool, and the
  GitHub MCP server's `web_search` tool (wire name
  `github-mcp-server-web_search`); the runtime also has an internal session
  option `webSearch` (not in the SDK's TypeScript types, visible in the
  bundled binary next to `toolSearch`). Order of attack for the builder:
  1. `excludedTools` with the three names above; verify in a session that
     a "what is in the news today" prompt produces no search.
  2. If hosted search still runs, also pass `webSearch: false` through the
     session config (cast past the types) and verify again.
  3. If neither works, the toggle becomes best effort: a sentence in the
     system prompt telling the agent not to use web search or fetch URLs,
     and the menu row says "(best effort)". The feature is not cut.
  The row is shown for every provider (BYOK sessions still have
  `web_fetch`). (My call.)
- After the session starts, the host builds the catalog: MCP servers
  through `client.rpc.mcp.discover` (returns name, type, source, enabled)
  and skills through `session.rpc.skills.list` (name, description, source,
  enabled, path). The catalog is sent in `ready` and again as `catalog`
  after every toggle. An item is shown on when it is enabled in the CLI's
  own list and not in our disabled list; an item the CLI itself has
  disabled is shown off and cannot be switched on from the pane.
- `mcp.setEnabled` / `skills.setEnabled` update `config.json`, call the
  session's live enable/disable RPC, and resend the catalog to every
  connected pane.
- `AgentSession.send(text, attachments)`:
  1. writes each file to `<workspaceDir>/attachments/<sanitized name>`;
  2. appends a footer to the prompt listing every attached file with its
     absolute path and MIME type, e.g.
     `Attached files (saved on disk):\n- C:\…\attachments\report.docx (application/vnd.openxmlformats-officedocument.wordprocessingml.document)`;
  3. passes images additionally as SDK `blob` attachments (base64 +
     mimeType) so a vision-capable model sees them. Documents are **not**
     passed as SDK `file` attachments (anti-goal: binary read into the
     prompt); the path in the footer is enough for officecli.
- The system prompt gets a paragraph telling the agent that attached files
  are under `attachments/` in the workspace, that Office files should be
  read and edited with officecli, and that an attached image can be placed
  in the document with `insertImage`.
- `maxPayload` on the WebSocket server is raised to cover the attachment
  limit.

### Task pane UI (`packages/ui`)

- Composer: a "+" icon button at the left of the input row; the send/stop
  button stays at the right. The menu is a popover anchored above the
  button, closed by click-outside or Escape. Layout follows the screenshot:
  icon + label rows, a check mark for the web search state, "›" on rows
  with a submenu. Submenus open in place of the menu (one level deep) with
  a back row at the top, which fits the narrow task pane better than a
  flyout. An empty list shows a short hint ("No MCP servers configured in
  Copilot CLI") above the "Open Copilot CLI settings…" row.
- Attachments state: `pending: PendingAttachment[]` (`File` + object URL
  for image preview). Chips render above the textarea. Drop target is the
  whole composer; while dragging over it the composer shows a dashed
  outline. Paste handler on the textarea picks image items from the
  clipboard.
- `ChatItem` of role `user` gains `attachments?: { name; mimeType; size; thumbnail?: string }[]`.
  Thumbnails are small data URLs (max 64 px) so local history stays small;
  the file bytes are not kept in history.
- A reusable `Toggle` component (role="switch", keyboard accessible,
  accent colour from `--accent`, works in dark mode) replaces every
  checkbox in the settings panel and is used in the submenus.
- The settings panel's info block shows the two Copilot CLI paths
  (MCP config file, skills folder) next to the existing workspace and
  allowed-directory lines.

### VSTO add-in (`vsto/OfficeCliAgent.Core`)

- The page ↔ add-in bridge (today only `{"__bridge":"ready"}`) gains
  `{"__bridge":"openPath","path"}`; the add-in opens the file or folder
  with the Windows default handler (`Process.Start` on the path, falling
  back to selecting it in Explorer). The relay must not forward bridge
  frames to the host. In the browser debugging page the row shows the path
  as text instead.
- Drag-and-drop of files into WebView2: `AllowExternalDrop` is true by
  default on the WinForms control; the builder verifies on the real Word
  task pane that a dropped file reaches the page's `drop` event and sets it
  explicitly if not.

### Verification on the real add-in

The builder uses the existing Debug-build workflow (see memory: Office
loads the repo Debug build) and checks in real Word: drop a png and a
docx, send "what is in these files", see the image described and officecli
called on the docx; toggle web search off, ask for today's news, see no
`web_search` call; see the CLI's MCP servers and the `officecli` skill from
`~/.copilot/skills` listed, switch one off, start a new chat and confirm it
is gone.

## Left to the builder

- Exact look of the menu, chips and toggle (sizes, icons: use inline SVG
  like the existing header buttons).
- How the "Searching the web…" tool row is keyed and finalized.
- Where `~/.copilot` is (honour `COPILOT_HOME` if set).
- Unit tests for the pure parts (name sanitizing, config helpers,
  attachment limits).
- Error texts for refused files.
