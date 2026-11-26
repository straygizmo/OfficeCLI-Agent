const LIST_TIMEOUT_MS = 15_000;

/**
 * The first model an OpenAI-compatible server lists (GET <baseUrl>/models), for a BYOK provider
 * saved without a model id. Local servers (Ollama, LM Studio, llama.cpp, vLLM …) answer this with
 * the models they have, so a local LLM works without typing its id. Throws with the reason otherwise.
 */
export async function firstServerModel(
  baseUrl: string,
  apiKey: string | undefined,
  headers: Record<string, string> | undefined,
  fetchImpl: typeof fetch = fetch,
): Promise<string> {
  const h: Record<string, string> = { ...(headers ?? {}) };
  if (apiKey) h.authorization = `Bearer ${apiKey}`;
  const url = `${baseUrl.replace(/\/+$/, "")}/models`;
  const response = await fetchImpl(url, { headers: h, signal: AbortSignal.timeout(LIST_TIMEOUT_MS) });
  if (!response.ok) throw new Error(`GET ${url} failed (HTTP ${response.status})`);
  const json = (await response.json()) as { data?: Array<{ id?: unknown }> };
  const id = json.data?.map((m) => m.id).find((v): v is string => typeof v === "string" && v.trim().length > 0);
  if (!id) throw new Error(`GET ${url} listed no models`);
  return id;
}
