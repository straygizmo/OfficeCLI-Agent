import { execFile } from "node:child_process";
import { defineTool, type Tool } from "@github/copilot-sdk";
import { z } from "zod";
import type { UiLocale } from "@officecli-agent/protocol";
import type { TaskpaneLink } from "./taskpane-link.js";

export interface OfficecliToolContext {
  officecliPath: string;
  locale: UiLocale;
  workspaceDir: string;
  link: TaskpaneLink;
  autoApproveReadOnly: () => boolean;
  permissionTimeoutMs: number;
}

const MAX_OUTPUT_CHARS = 80_000;
const TIMEOUT_MS = 180_000;

/** officecli verbs that never modify a document. */
const READ_ONLY_VERBS = new Set(["view", "get", "query", "help", "validate", "load_skill", "plugins", "--help", "--version"]);

/**
 * OfficeCLI as a host-side tool.
 *
 * We deliberately do not use `officecli mcp`: on Windows, `view … screenshot` and
 * `stats` default to `--render auto`, which drives the *installed* PowerPoint/Word
 * through COM. Attaching to the running instance that hosts this add-in crashes
 * PowerPoint. Running the CLI ourselves lets us force the built-in HTML renderer,
 * auto-approve read-only verbs, and ask the user before anything that writes.
 */
export function createOfficecliTool(ctx: OfficecliToolContext): Tool<unknown> {
  return defineTool("officecli", {
    description:
      "Run an OfficeCLI command against .docx/.xlsx/.pptx files on disk (high-fidelity read/analyze/edit, no Office needed). " +
      "Pass the raw command line after `officecli`, as a string or argv array. Always use absolute paths. Examples: " +
      "`view report.docx outline`, `view deck.pptx issues --json`, `view deck.pptx screenshot -o out.png --page 1`, " +
      "`get budget.xlsx /Sheet1 --depth 1 --json`, `query report.docx \"paragraph[style=Heading1]\" --json`, " +
      "`set deck.pptx /slide[1]/shape[1] --prop text=\"Hello\"`, `add deck.pptx / --type slide --prop title=\"Q4\"`, " +
      "`batch file.docx --input ops.json`, `merge tpl.docx out.docx --data '{...}'`. " +
      "Run `help <docx|xlsx|pptx> <set|add|query>` when unsure about properties. " +
      "Read-only verbs (view/get/query/help/validate) run immediately; writes ask the user for approval. " +
      "Screenshots/stats are rendered with the built-in engine (never the running Office app).",
    parameters: z.object({
      command: z.union([z.string(), z.array(z.string())]).describe("Command line without the leading `officecli`."),
    }),
    skipPermission: true,
    handler: async ({ command }) => {
      let argv = typeof command === "string" ? splitCommandLine(command) : command.map(String);
      if (argv.length === 0) return { textResultForLlm: "Empty command.", resultType: "failure" as const };
      argv = forceHtmlRenderer(argv);

      const verb = argv[0].toLowerCase();
      if (!READ_ONLY_VERBS.has(verb) || !ctx.autoApproveReadOnly()) {
        const approved = await ctx.link.askPermission(
          { kind: "officecli", title: ctx.locale === "ja" ? "officecli を実行します" : "Run officecli", detail: "officecli " + argv.map(quoteForDisplay).join(" ") },
          ctx.permissionTimeoutMs,
        );
        if (!approved) return { textResultForLlm: "The user declined to run this officecli command.", resultType: "failure" as const };
      }

      const { code, stdout, stderr } = await run(ctx.officecliPath, argv, ctx.workspaceDir);
      const out = [stdout.trim(), stderr.trim() ? `[stderr]\n${stderr.trim()}` : ""].filter(Boolean).join("\n");
      const text = clip(out || "(no output)");
      if (code !== 0) {
        return { textResultForLlm: `officecli exited with code ${code}.\n${text}`, resultType: "failure" as const };
      }
      return text;
    },
  }) as Tool<unknown>;
}

/** Make `view … screenshot|stats` use the HTML pipeline unless the caller chose explicitly. */
function forceHtmlRenderer(argv: string[]): string[] {
  const isView = argv[0]?.toLowerCase() === "view";
  const mode = argv[2]?.toLowerCase();
  if (!isView || (mode !== "screenshot" && mode !== "stats")) return argv;
  const idx = argv.findIndex((a) => a === "--render" || a.startsWith("--render="));
  if (idx === -1) return [...argv, "--render", "html"];
  // Never allow the native backend: it would attach to the running Office instance.
  const next = [...argv];
  if (next[idx] === "--render") next[idx + 1] = "html";
  else next[idx] = "--render=html";
  return next;
}

function run(bin: string, argv: string[], cwd: string): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    execFile(
      bin,
      argv,
      { cwd, timeout: TIMEOUT_MS, maxBuffer: 32 * 1024 * 1024, windowsHide: true, encoding: "utf8" },
      (err, stdout, stderr) => {
        const code = err ? ((err as NodeJS.ErrnoException & { code?: number | string }).code as number | undefined) ?? 1 : 0;
        resolve({ code: typeof code === "number" ? code : 1, stdout: String(stdout ?? ""), stderr: String(stderr ?? "") + (err && typeof code !== "number" ? `\n${err.message}` : "") });
      },
    );
  });
}

function clip(text: string): string {
  return text.length > MAX_OUTPUT_CHARS ? text.slice(0, MAX_OUTPUT_CHARS) + `\n…(truncated, ${text.length} chars total)` : text;
}

function quoteForDisplay(arg: string): string {
  return /[\s"]/.test(arg) ? `"${arg.replace(/"/g, '\\"')}"` : arg;
}

/** Minimal shell-style tokenizer (double/single quotes, backslash-escaped quotes inside double quotes). */
export function splitCommandLine(line: string): string[] {
  const out: string[] = [];
  let cur = "";
  let quote: '"' | "'" | null = null;
  let has = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (quote) {
      if (ch === "\\" && quote === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === quote) quote = null;
      else cur += ch;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      has = true;
    } else if (/\s/.test(ch)) {
      if (has || cur) out.push(cur);
      cur = "";
      has = false;
    } else {
      cur += ch;
    }
  }
  if (has || cur) out.push(cur);
  return out;
}
