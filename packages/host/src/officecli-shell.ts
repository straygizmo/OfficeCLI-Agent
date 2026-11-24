import type { PermissionRequest } from "@microsoft/copilot-sdk";

type ShellRequest = Extract<PermissionRequest, { kind: "shell" }>;

/** Commands that may feed data into `officecli` through a pipe (they only print or read). */
const FEEDERS = new Set(["echo", "write-output", "printf", "cat", "type", "get-content", "gc"]);

/**
 * True when a shell permission request does nothing but run officecli, optionally with a
 * printing/reading command piped into it (e.g. `echo '[…]' | officecli batch x.xlsx`).
 *
 * Used to honour "auto-approve OfficeCLI" when the model reaches officecli through the shell
 * instead of the officecli tool. Deliberately strict: anything it cannot fully account for
 * (command chaining, redirection, substitution, URLs, other programs, the native renderer)
 * returns false and the user is asked as usual.
 */
export function isOfficecliOnlyShell(request: ShellRequest): boolean {
  if (request.managedApprovalRequired || request.requestSandboxBypass) return false;
  if (request.hasWriteFileRedirection || (request.possibleUrls?.length ?? 0) > 0) return false;
  // `request.commands` is not relied on: SDK 1.x reported a whole PowerShell command line as a single
  // identifier and left hasWriteFileRedirection false even with `>`. Our own parse decides.

  const segments = splitPipeline(request.fullCommandText);
  if (!segments || segments.length === 0) return false;
  for (let i = 0; i < segments.length; i++) {
    const argv = segments[i];
    if (argv.length === 0) return false;
    if (isOfficecliName(argv[0])) {
      if (usesNativeRenderer(argv.slice(1))) return false;
      continue;
    }
    // Only the head of a pipeline may be a feeder, and it must actually feed officecli.
    if (i === 0 && segments.length > 1 && FEEDERS.has(baseName(argv[0]))) continue;
    return false;
  }
  return true;
}

/** `view … screenshot|stats` without `--render html` attaches to the running Office app (crash). */
export function usesNativeRenderer(argv: string[]): boolean {
  if (argv[0]?.toLowerCase() !== "view") return false;
  const mode = argv[2]?.toLowerCase();
  if (mode !== "screenshot" && mode !== "stats") return false;
  const idx = argv.findIndex((a) => a === "--render" || a.startsWith("--render="));
  if (idx === -1) return true;
  const value = argv[idx] === "--render" ? argv[idx + 1] : argv[idx].slice("--render=".length);
  return value?.toLowerCase() !== "html";
}

function baseName(identifier: string): string {
  const name = identifier.replace(/^["']|["']$/g, "").split(/[\\/]/).pop() ?? "";
  return name.toLowerCase();
}

function isOfficecliName(identifier: string): boolean {
  return /^officecli(\.exe)?$/.test(baseName(identifier));
}

/**
 * Split a command line into pipeline segments of argv tokens (quotes removed).
 * Returns null for anything beyond plain words, quotes and single pipes: separators, `&`,
 * redirections, grouping, substitutions, backticks or line breaks outside quotes.
 */
export function splitPipeline(text: string): string[][] | null {
  // Substitutions expand even inside double quotes (bash `$(…)`, PowerShell `$(…)`/backtick escapes).
  if (/`|\$\(|\$\{|<\(|>\(/.test(text)) return null;
  // bash escapes `\"` inside double quotes while PowerShell does not: the two could disagree on
  // where a string ends, so do not try to decide.
  if (/\\["']/.test(text)) return null;
  const segments: string[][] = [];
  let argv: string[] = [];
  let cur = "";
  let has = false;
  let quote: '"' | "'" | null = null;
  const endWord = () => {
    if (has || cur) argv.push(cur);
    cur = "";
    has = false;
  };
  const trimmed = text.trim();
  for (let i = 0; i < trimmed.length; i++) {
    const ch = trimmed[i];
    if (quote) {
      if (ch === quote) {
        // PowerShell doubles a quote to escape it inside a string of the same kind.
        if (trimmed[i + 1] === quote) {
          cur += ch;
          i++;
        } else quote = null;
      } else cur += ch;
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      has = true;
    } else if (ch === "|") {
      endWord();
      if (argv.length === 0) return null; // `||` or a leading pipe
      segments.push(argv);
      argv = [];
    } else if (ch === " " || ch === "\t") {
      endWord();
    } else if (/[;&<>(){}\r\n]/.test(ch)) {
      return null;
    } else {
      cur += ch;
    }
  }
  if (quote) return null;
  endWord();
  if (argv.length === 0) return null; // trailing pipe or empty command
  segments.push(argv);
  return segments;
}
