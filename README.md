# officecli-agent

English | [日本語](README-ja.md)

An open-source AI agent **VSTO add-in** for Word / Excel / PowerPoint (Windows desktop).

- **Agent**: [GitHub Copilot SDK](https://github.com/github/copilot-sdk) (the agent runtime of Copilot CLI). Works with either a Copilot subscription (CLI sign-in) or **BYOK** (your own keys for OpenAI / Azure OpenAI / Anthropic / Ollama, etc.). There is no custom OAuth implementation.
- **Document editing**: [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) is wired in as a host-side tool (the host runs the CLI directly). It enables high-fidelity OOXML-level editing: pivots, tracked changes, animations, equations, Mermaid diagrams, template merges, issue diagnostics, screenshots and more.
- **Live editing**: Reads and writes the *currently open* document directly through the Office COM object model (undoable).
- **Importing other files**: The agent reads `.docx` / `.xlsx` / `.pptx` files on your PC and inserts them into the open document (Word: insert into the body/selection, PowerPoint: insert slides, Excel: add sheets).
- **Inserting pictures**: `office_insert_image` puts PNG / JPEG / GIF / BMP / TIFF / SVG / EMF / WMF files into the open document (Word: at the selection, Excel: at a cell, PowerPoint: on a slide).

The task pane and ribbon follow the Office UI language: Japanese in Japanese environments, English everywhere else.

## How it works

```
┌─ Office (Word / Excel / PowerPoint) ─────────────────────────────────┐
│  VSTO add-in (vsto/)                                                  │
│   ├ Ribbon: "Agent" toggle on the Home tab                            │
│   ├ Task pane: React chat UI (packages/ui) hosted in WebView2         │
│   ├ HostRelay: WebSocket relay between the UI and the host            │
│   └ *Ops.cs: runs office.request via COM (read/write/export/import)   │
└───────────────────────────┬──────────────────────────────────────────┘
                            │ ws://127.0.0.1:4320 (auto-started)
┌─ Local host (packages/host, Node) ────────────────────────────────────┐
│  GitHub Copilot SDK → Copilot CLI (agent loop, Copilot or BYOK)        │
│   ├ Custom tools: office_get_document_context / office_export_document │
│   │               / office_import_file / word_* / excel_* / ppt_*      │
│   │               / office_insert_image                                │
│   ├ officecli tool: runs the CLI directly (reads auto-approved,        │
│   │                 writes need confirmation)                          │
│   └ Built-in: file view/search/edit, shell (approved in the task pane) │
└───────────────────────────────────────────────────────────────────────┘
```

OfficeCLI only works on files on disk, so for the open document the add-in writes out a copy, OfficeCLI edits it, and the add-in imports it back (a round trip) to apply the changes live.

| Goal | Path the agent takes |
|---|---|
| Write cells, insert/replace text, slide text | `excel_set_range` / `word_insert_text` / `ppt_set_shape_text`, etc. (COM, applied instantly, undoable) |
| Bring in content from another Office file | Inspect with `officecli view/get` → `office_import_file` |
| Pivots, tracked changes, animations, equations, diagrams, etc. | `office_export_document` → edit with `officecli` → `office_import_file` |
| Diagnose document issues / check appearance | `office_export_document` → `officecli view … issues / screenshot` |

For Excel only, existing sheets are never replaced: the agent either adds a sheet or applies results computed with officecli to cells via `excel_set_range`. This avoids breaking references and losing Undo.

## Requirements (Windows)

- Microsoft 365 / Office 2021 or later for Windows
- Visual Studio 2022 with the "Office/SharePoint development" workload (to build the VSTO add-ins)
- Microsoft Edge WebView2 Runtime
- Node.js 20.19 or later, pnpm 9
- [GitHub Copilot CLI](https://github.com/github/copilot-cli) (`copilot` on PATH). To use Copilot models, run `copilot` once and sign in. **Not needed for BYOK only.**
- [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) (`officecli` on PATH: `irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex`)

## Setup

```powershell
pnpm install
pnpm build          # protocol → host → ui
```

Then open `vsto\OfficeCliAgent.<App>\OfficeCliAgent.<App>.sln` in Visual Studio and press F5 (see [vsto/README.md](vsto/README.md) for details). The manifest signing certificates (`*_TemporaryKey.pfx`) are not included in the repository, so the first time you build, create one in each project's Properties → Signing → "Create Test Certificate...". The add-in starts the host (Node) automatically, so you don't need to launch it separately.

You can also run the host by hand and try the chat UI in a browser:

```powershell
pnpm start          # http://localhost:4320/taskpane.html?hostApp=Word  (Office operations will fail)
```

Add `&lang=ja` or `&lang=en` to the URL to force the UI language; otherwise the browser language is used.

## Configuration

### Host: `%USERPROFILE%\.officecli-agent\config.json` (created on first run)

```json
{
  "port": 4320,
  "defaultProvider": "copilot",
  "model": "claude-sonnet-4.5",
  "providers": [],
  "allowedDirectories": ["C:\\Users\\you\\Documents"],
  "officecliPath": "officecli",
  "autoApproveReadOnly": true,
  "permissionTimeoutSeconds": 180
}
```

- `allowedDirectories`: Folders the agent may read and write without asking. Reading a file elsewhere with `office_import_file` shows an approval prompt in the task pane.
- `autoApproveReadOnly`: Auto-approve read-only tool calls (file viewing and the read-only OfficeCLI commands view/get/query/help/validate). Shell commands, file writes and URL access always ask for approval in the task pane.
- `model`: Copilot model ID (defaults to the Copilot CLI default). Can also be chosen in the task pane settings (takes effect from the next session).
- `copilotCliPath`: Path to the executable when `copilot` is not on PATH.

### Add-in: `%LOCALAPPDATA%\OfficeCliAgent\addin.json`

Port, path to `node`, host/UI folders, and whether to auto-start the host. When omitted, the repository is located from the add-in DLL's location. See [vsto/README.md](vsto/README.md) for details.

## BYOK (your own API keys / local models)

Register providers from the task pane under "Settings → BYOK providers → Add" (presets: OpenAI / Anthropic / Azure OpenAI / Ollama / LM Studio). They are saved in the host's `config.json`, and API keys are never sent back to the task pane. You can also edit the file directly.

```json
{
  "defaultProvider": "ollama",
  "providers": [
    { "name": "openai", "type": "openai", "baseUrl": "https://api.openai.com/v1", "apiKeyEnv": "OPENAI_API_KEY", "model": "gpt-6-luna" },
    { "name": "anthropic", "type": "anthropic", "baseUrl": "https://api.anthropic.com", "apiKey": "sk-ant-…", "model": "claude-sonnet-5.5" },
    { "name": "azure", "type": "azure", "baseUrl": "https://<resource>.openai.azure.com", "apiKey": "…", "model": "gpt-6-luna", "azureApiVersion": "2024-10-21" },
    { "name": "ollama", "type": "openai", "baseUrl": "http://localhost:11434/v1", "model": "llama3.1", "modelProvider": "ollama" }
  ]
}
```

- `type`: `openai` (any OpenAI-compatible server: OpenAI / Ollama / vLLM / LiteLLM / LM Studio / Foundry Local), `azure` (native Azure OpenAI), or `anthropic`.
- `wireApi`: Defaults to `responses` for OpenAI (`api.openai.com`) and Azure OpenAI (gpt-5.6 or later assumed), and `completions` for other OpenAI-compatible servers. Set it explicitly if needed. For Azure, `model` is the deployment name.
- Use `apiKeyEnv` instead of `apiKey` to read the key from an environment variable. Local models need no key.
- BYOK requires a model ID (the provider's default `model`, or one entered in the task pane).
- With BYOK, usage and rate limits are those of the provider; no Copilot premium requests are consumed. Pick a model that supports tool calling (for Ollama, Llama 3.1 or later, Qwen 2.5, etc.).

## Example requests

- "Summarize the structure of this document"
- "Add Sheet1 of `C:\Users\me\Documents\Sales.xlsx` to the end of this workbook"
- "Insert slides 3–5 of `Proposal_old.pptx` after slide 2 of this presentation"
- "Put the sum of B2:B20 in B21 and format it as currency"
- "Diagnose and fix layout problems (such as overflowing text) in this presentation"

## Repository layout

```
packages/protocol   Message types between add-in and host (WebSocket), including OfficeOp
packages/host       Local host (Node): Copilot SDK session, OfficeCLI tool, Office tools, BYOK
packages/ui         Chat UI (Vite + React). Works over WebView2 postMessage and over WebSocket in a browser
vsto/               VSTO add-ins: shared library (Core) and Word / Excel / PowerPoint solutions
```

## Design note: OfficeCLI and a running Office

By default (`--render auto`), OfficeCLI's `view … screenshot` / `stats` render by driving the Word / PowerPoint installed on Windows through COM. Attaching to the Office instance that hosts the add-in from an external process crashes PowerPoint, so the host runs OfficeCLI as its own tool rather than as an MCP server and always adds `--render html` (the built-in renderer) to these commands. The system prompt also forbids driving Office through COM from the shell (PowerShell, etc.).

## Security notes

- The agent can both read local files and access the web (built into Copilot CLI). The system prompt tells it to treat instructions embedded in file or web content as data, but always review shell commands, writes and URL access before approving them in the task pane.
- The host binds only to `127.0.0.1` and has no authentication. Other processes on the same PC can connect to it.
- Not supported on Office on the web / Mac / iPad (requires VSTO and a local host).

## Credits

This project is built on the following open-source software:

- **[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)** (iOfficeAI, Apache-2.0) — the document engine that reads, writes, renders and diagnoses Word / Excel / PowerPoint files. All of this add-in's advanced editing features are provided by calling the OfficeCLI CLI from the host.
- **[GitHub Copilot SDK](https://github.com/github/copilot-sdk)** (GitHub, MIT) — the agent runtime. The tool-calling loop, MCP connections, BYOK and permission control are features of the Copilot SDK / Copilot CLI.
- [office-agents](https://github.com/hewliyang/office-agents) (hewliyang, MIT) — inspiration for embedding a BYOK agent in an Office add-in.
- [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) — hosts the task pane UI.

OfficeCLI and the GitHub Copilot SDK must be installed separately and are not bundled with this repository. Follow each project's license and terms of use.

## License

The code in this repository is licensed under the MIT License.
