import type { UiLocale } from "@officecli-agent/protocol";

const PROBE_TIMEOUT_MS = 8_000;

/** The parts of a BYOK provider needed to ask it why it rejected a request. */
export interface ProbeTarget {
  /** The provider's name as shown in the task pane. */
  name: string;
  type: "openai" | "azure" | "anthropic";
  baseUrl: string;
  apiKey?: string;
  headers?: Record<string, string>;
}

/** The fields of a Copilot SDK `session.error` event this module looks at. */
export interface SessionErrorLike {
  message: string;
  errorType?: string;
  statusCode?: number;
  providerErrorCode?: string;
}

/** True for an error that says the provider rejected the credentials (HTTP 401/403). */
export function isAuthError(error: SessionErrorLike): boolean {
  return (
    error.errorType === "authentication" ||
    error.errorType === "authorization" ||
    error.statusCode === 401 ||
    error.statusCode === 403 ||
    /authentication failed/i.test(error.message)
  );
}

function trimSlash(url: string): string {
  return url.replace(/\/+$/, "");
}

/** The provider's model list endpoint and the headers that authenticate a request to it. */
export function modelsRequest(target: ProbeTarget): { url: string; headers: Record<string, string> } {
  const base = trimSlash(target.baseUrl);
  const headers: Record<string, string> = { ...(target.headers ?? {}) };
  switch (target.type) {
    case "anthropic":
      if (target.apiKey) headers["x-api-key"] = target.apiKey;
      headers["anthropic-version"] = "2023-06-01";
      return { url: /\/v1$/.test(base) ? `${base}/models` : `${base}/v1/models`, headers };
    case "azure":
      if (target.apiKey) headers["api-key"] = target.apiKey;
      // The base URL may be the resource (https://x.openai.azure.com) or its /openai/v1 endpoint.
      return { url: `${base.replace(/\/openai(\/v1)?$/, "")}/openai/v1/models`, headers };
    default:
      if (target.apiKey) headers.authorization = `Bearer ${target.apiKey}`;
      return { url: `${base}/models`, headers };
  }
}

/** The reason in an OpenAI / Anthropic / Azure style error body ({ error: { message, code } }), if any. */
export function errorReason(body: string): string | undefined {
  try {
    const json = JSON.parse(body) as { error?: { message?: unknown; code?: unknown } | string; message?: unknown };
    const err = json.error;
    if (typeof err === "string" && err.trim()) return err.trim();
    if (err && typeof err === "object") {
      const message = typeof err.message === "string" ? err.message.trim() : "";
      const code = typeof err.code === "string" ? err.code.trim() : "";
      if (message) return code ? `${message} (${code})` : message;
      if (code) return code;
    }
    if (typeof json.message === "string" && json.message.trim()) return json.message.trim();
  } catch {
    // Not JSON: fall through.
  }
  const text = body.trim();
  return text && text.length <= 300 && !text.startsWith("<") ? text : undefined;
}

/**
 * Ask the provider directly (GET its model list with the same key) why it rejects the key: the
 * runtime's own message only says "check COPILOT_PROVIDER_API_KEY", which is not what the user set.
 * Returns the status and the provider's reason, or undefined when the probe itself fails or succeeds.
 */
export async function probeAuthFailure(
  target: ProbeTarget,
  fetchImpl: typeof fetch = fetch,
): Promise<{ status: number; reason?: string } | undefined> {
  const { url, headers } = modelsRequest(target);
  try {
    const response = await fetchImpl(url, { headers, signal: AbortSignal.timeout(PROBE_TIMEOUT_MS) });
    if (response.ok) return undefined;
    return { status: response.status, reason: errorReason(await response.text()) };
  } catch {
    return undefined;
  }
}

/** A task-pane message for a BYOK provider that rejected the API key. */
export function describeAuthError(
  locale: UiLocale,
  target: ProbeTarget,
  error: SessionErrorLike,
  probe: { status: number; reason?: string } | undefined,
): string {
  const ja = locale === "ja";
  const status = probe?.status ?? error.statusCode;
  const reason = probe?.reason ?? error.providerErrorCode;
  const head = !target.apiKey
    ? ja
      ? `BYOK プロバイダー「${target.name}」に API キーが設定されていません。`
      : `The BYOK provider '${target.name}' has no API key.`
    : ja
      ? `BYOK プロバイダー「${target.name}」が API キーを受け付けませんでした${status ? ` (HTTP ${status})` : ""}。`
      : `The BYOK provider '${target.name}' rejected the API key${status ? ` (HTTP ${status})` : ""}.`;
  const why = reason ? (ja ? `\n理由: ${reason}` : `\nReason: ${reason}`) : "";
  const fix = ja
    ? `\n設定 → BYOK プロバイダー → 「${target.name}」の編集で API キーを確認・更新してから、新しいチャットを始めてください。`
    : `\nCheck or replace the key under Settings → BYOK providers → '${target.name}' (Edit), then start a new chat.`;
  return head + why + fix;
}
