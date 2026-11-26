import fs from "node:fs/promises";
import path from "node:path";
import { resolveApiKey, resolveBaseUrl, type StoredImageProvider } from "./config.js";

export type ImageAspect = "square" | "landscape" | "portrait";

/** The API version Azure OpenAI image deployments are served under when none is configured. */
export const DEFAULT_AZURE_IMAGE_API_VERSION = "2025-04-01-preview";
/** Image models can take a minute or more per picture. */
const REQUEST_TIMEOUT_MS = 300_000;
const LIST_TIMEOUT_MS = 30_000;

export interface ImageRequest {
  url: string;
  headers: Record<string, string>;
  body: Record<string, unknown>;
}

export interface GeneratedImage {
  bytes: Buffer;
  /** File extension without the dot, from the image's magic bytes ("png" when unrecognized). */
  ext: string;
}

/**
 * Pixel size for an aspect. The wide side is 1536 for the gpt-image family and 1792 for dall-e-3
 * (recognized by model id; an Azure deployment with another name gets the gpt-image sizes).
 */
export function imageSize(model: string, aspect: ImageAspect): string {
  if (aspect === "square") return "1024x1024";
  const wide = /^dall-e-3/i.test(model) ? 1792 : 1536;
  return aspect === "landscape" ? `${wide}x1024` : `1024x${wide}`;
}

/** Aspect ratios as the Gemini API names them. */
const GEMINI_ASPECT: Record<ImageAspect, string> = { square: "1:1", landscape: "3:2", portrait: "2:3" };

/**
 * The HTTP request for one picture. Throws when the provider is missing a base URL, or a model
 * (only the OpenAI-compatible type may go without one: the server then uses its own default).
 */
export function buildImageRequest(provider: StoredImageProvider, prompt: string, aspect: ImageAspect): ImageRequest {
  const model = provider.model?.trim() ?? "";
  if (!model && provider.type !== "openai") throw new Error("The image generation provider has no model id.");
  const baseUrl = resolveBaseUrl(provider)?.replace(/\/+$/, "");
  if (!baseUrl) {
    throw new Error(
      provider.resourceNameEnv
        ? `The image generation provider has no base URL and the environment variable ${provider.resourceNameEnv} is not set (restart Office after setting it).`
        : "The image generation provider has no base URL.",
    );
  }
  const apiKey = resolveApiKey(provider);
  const headers: Record<string, string> = { "content-type": "application/json", ...(provider.headers ?? {}) };
  // No response_format: the gpt-image models reject it (they always answer with base64); a URL answer is downloaded instead.
  const body: Record<string, unknown> = { prompt, n: 1, size: imageSize(model, aspect) };

  if (provider.type === "gemini") {
    if (apiKey) headers["x-goog-api-key"] = apiKey;
    return {
      url: `${baseUrl}/models/${encodeURIComponent(model)}:generateContent`,
      headers,
      body: {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { responseModalities: ["IMAGE"], imageConfig: { aspectRatio: GEMINI_ASPECT[aspect] } },
      },
    };
  }
  if (provider.type === "azure") {
    if (apiKey) headers["api-key"] = apiKey;
    const version = provider.azureApiVersion?.trim() || DEFAULT_AZURE_IMAGE_API_VERSION;
    return {
      url: `${baseUrl}/openai/deployments/${encodeURIComponent(model)}/images/generations?api-version=${encodeURIComponent(version)}`,
      headers,
      body,
    };
  }
  if (apiKey) headers.authorization = `Bearer ${apiKey}`;
  return { url: `${baseUrl}/images/generations`, headers, body: model ? { model, ...body } : body };
}

/** File extension for image bytes, from their signature. */
export function sniffImageExt(bytes: Buffer): string {
  if (bytes.length >= 4 && bytes[0] === 0x89 && bytes.toString("latin1", 1, 4) === "PNG") return "png";
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "jpg";
  if (bytes.length >= 12 && bytes.toString("latin1", 0, 4) === "RIFF" && bytes.toString("latin1", 8, 12) === "WEBP") return "webp";
  if (bytes.length >= 4 && bytes.toString("latin1", 0, 4) === "GIF8") return "gif";
  return "png";
}

async function errorText(response: Response): Promise<string> {
  const text = await response.text().catch(() => "");
  try {
    const message = (JSON.parse(text) as { error?: { message?: string } | string }).error;
    if (typeof message === "string") return message;
    if (message?.message) return message.message;
  } catch {
    /* not JSON: use the raw text */
  }
  return text.slice(0, 500) || response.statusText;
}

interface GeminiResponse {
  candidates?: Array<{
    finishReason?: string;
    content?: { parts?: Array<{ text?: string; thought?: boolean; inlineData?: { mimeType?: string; data?: string } }> };
  }>;
  promptFeedback?: { blockReason?: string };
}

/** The final picture of a generateContent answer: the last image part that is not an interim "thought" image. */
export function geminiImageBytes(json: GeminiResponse): Buffer {
  const candidate = json.candidates?.[0];
  const images = (candidate?.content?.parts ?? []).filter((p) => p.inlineData?.data);
  const final = images.filter((p) => !p.thought).pop() ?? images.pop();
  if (final) return Buffer.from(final.inlineData!.data!, "base64");
  // No picture: the model answered in words (often a refusal) or the prompt was blocked.
  const said = (candidate?.content?.parts ?? []).map((p) => p.text ?? "").join(" ").trim().slice(0, 500);
  const reason = json.promptFeedback?.blockReason ?? candidate?.finishReason;
  throw new Error(`The image generation provider returned no image${reason ? ` (${reason})` : ""}${said ? `: ${said}` : "."}`);
}

/** Ask the provider for one picture and return its bytes. */
export async function generateImage(
  provider: StoredImageProvider,
  prompt: string,
  aspect: ImageAspect,
  fetchImpl: typeof fetch = fetch,
): Promise<GeneratedImage> {
  const request = buildImageRequest(provider, prompt, aspect);
  const response = await fetchImpl(request.url, {
    method: "POST",
    headers: request.headers,
    body: JSON.stringify(request.body),
    signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
  });
  if (!response.ok) throw new Error(`Image generation failed (HTTP ${response.status}): ${await errorText(response)}`);

  const json = (await response.json()) as { data?: Array<{ b64_json?: string; url?: string }> };
  const first = json.data?.[0];
  let bytes: Buffer;
  if (provider.type === "gemini") {
    bytes = geminiImageBytes(json as GeminiResponse);
  } else if (first?.b64_json) {
    bytes = Buffer.from(first.b64_json, "base64");
  } else if (first?.url) {
    const download = await fetchImpl(first.url, { signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
    if (!download.ok) throw new Error(`Downloading the generated image failed (HTTP ${download.status}).`);
    bytes = Buffer.from(await download.arrayBuffer());
  } else {
    throw new Error("The image generation provider returned no image.");
  }
  if (bytes.length === 0) throw new Error("The image generation provider returned an empty image.");
  return { bytes, ext: sniffImageExt(bytes) };
}

/** Image models from an OpenAI-style model list, newest first; every model when none looks like an image model. */
export function pickOpenAiImageModels(data: Array<{ id?: string; created?: number }>): string[] {
  const all = data.filter((m): m is { id: string; created?: number } => typeof m.id === "string" && m.id.length > 0);
  // The list has no capability field: go by the id (gpt-image-*, chatgpt-image-*, dall-e-*, and similar on other servers).
  const images = all.filter((m) => /image|dall-e/i.test(m.id));
  return (images.length ? images : all).sort((a, b) => (b.created ?? 0) - (a.created ?? 0) || a.id.localeCompare(b.id)).map((m) => m.id);
}

/** Image models from a Gemini model list: the ones generateContent can call (Imagen uses another method). */
export function pickGeminiImageModels(models: Array<{ name?: string; supportedGenerationMethods?: string[] }>): string[] {
  const callable = models
    .filter((m) => typeof m.name === "string" && (m.supportedGenerationMethods ?? []).includes("generateContent"))
    .map((m) => m.name!.replace(/^models\//, ""));
  const images = callable.filter((id) => /image|banana/i.test(id));
  return images.length ? images : callable;
}

/** Ask the provider which image models it offers. Throws with the reason when it cannot say. */
export async function listImageModels(provider: StoredImageProvider, fetchImpl: typeof fetch = fetch): Promise<string[]> {
  if (provider.type === "azure") throw new Error("Azure OpenAI deployments cannot be listed with an API key; enter the deployment name.");
  const baseUrl = resolveBaseUrl(provider)?.replace(/\/+$/, "");
  if (!baseUrl) throw new Error("Enter a base URL first.");
  const apiKey = resolveApiKey(provider);
  const headers: Record<string, string> = { ...(provider.headers ?? {}) };
  const gemini = provider.type === "gemini";
  if (apiKey) headers[gemini ? "x-goog-api-key" : "authorization"] = gemini ? apiKey : `Bearer ${apiKey}`;
  const response = await fetchImpl(gemini ? `${baseUrl}/models?pageSize=1000` : `${baseUrl}/models`, { headers, signal: AbortSignal.timeout(LIST_TIMEOUT_MS) });
  if (!response.ok) throw new Error(`HTTP ${response.status}: ${await errorText(response)}`);
  const json = (await response.json()) as { data?: Array<{ id?: string; created?: number }>; models?: Array<{ name?: string; supportedGenerationMethods?: string[] }> };
  return gemini ? pickGeminiImageModels(json.models ?? []) : pickOpenAiImageModels(json.data ?? []);
}

/** Write `bytes` as <dir>/<baseName>.<ext>, adding -2, -3… rather than overwriting. Returns the path. */
export async function saveGeneratedImage(dir: string, baseName: string, image: GeneratedImage): Promise<string> {
  await fs.mkdir(dir, { recursive: true });
  for (let n = 1; ; n++) {
    const file = path.join(dir, `${baseName}${n === 1 ? "" : `-${n}`}.${image.ext}`);
    try {
      await fs.writeFile(file, image.bytes, { flag: "wx" });
      return file;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err;
    }
  }
}
