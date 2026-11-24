/**
 * Tool arguments are not validated against the schema before the handler runs, and models sometimes
 * call word_replace_text with other names for the two texts (seen: `text` / `replaceWith`). Accept the
 * common synonyms instead of failing the call and making the model retry.
 */
const FIND_ALIASES = ["find", "findText", "search", "searchText", "text", "oldText", "old", "query", "pattern"];
const REPLACE_ALIASES = ["replace", "replaceWith", "replacement", "replaceText", "newText", "new", "with"];

export interface ReplaceTextArgs {
  find: string;
  replace: string;
  matchCase: boolean;
  replaceAll: boolean;
}

export function normalizeReplaceArgs(args: unknown): ReplaceTextArgs {
  const a = (args && typeof args === "object" ? args : {}) as Record<string, unknown>;
  const pick = (names: string[]): string | undefined => {
    for (const n of names) if (typeof a[n] === "string") return a[n] as string;
    return undefined;
  };
  const find = pick(FIND_ALIASES);
  if (!find) throw new Error("word_replace_text needs `find` (the text to search for) and `replace` (the text to put in its place).");
  return {
    find,
    replace: pick(REPLACE_ALIASES) ?? "",
    matchCase: a.matchCase === true,
    replaceAll: a.replaceAll !== false,
  };
}
