import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { imageProviderForListing, mergeImageProvider, type StoredImageProvider } from "../src/config.js";
import {
  buildImageRequest,
  geminiImageBytes,
  generateImage,
  imageSize,
  listImageModels,
  pickGeminiImageModels,
  pickOpenAiImageModels,
  saveGeneratedImage,
  sniffImageExt,
} from "../src/image-gen.js";

const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const JPG = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]);
const openai: StoredImageProvider = { type: "openai", baseUrl: "https://api.openai.com/v1/", model: "gpt-image-2.6-flare", apiKey: "sk-test" };

function fakeFetch(responses: Response[], calls: Array<{ url: string; init?: RequestInit }> = []): typeof fetch {
  return (async (url: string | URL | Request, init?: RequestInit) => {
    calls.push({ url: String(url), init });
    return responses.shift()!;
  }) as typeof fetch;
}

test("imageSize picks the wide side by model family", () => {
  assert.equal(imageSize("gpt-image-2.6-flare", "square"), "1024x1024");
  assert.equal(imageSize("gpt-image-2.6-flare", "landscape"), "1536x1024");
  assert.equal(imageSize("gpt-image-2.6-flare", "portrait"), "1024x1536");
  assert.equal(imageSize("dall-e-3", "landscape"), "1792x1024");
});

test("buildImageRequest: OpenAI-compatible endpoint with a bearer key", () => {
  const req = buildImageRequest(openai, "a cat", "square");
  assert.equal(req.url, "https://api.openai.com/v1/images/generations");
  assert.equal(req.headers.authorization, "Bearer sk-test");
  assert.deepEqual(req.body, { model: "gpt-image-2.6-flare", prompt: "a cat", n: 1, size: "1024x1024" });
});

test("buildImageRequest: no key means no authorization header (local servers)", () => {
  const req = buildImageRequest({ type: "openai", baseUrl: "http://localhost:8080/v1", model: "sd" }, "x", "square");
  assert.equal(req.headers.authorization, undefined);
});

test("buildImageRequest: Azure deployment URL, api-key header, resource name from the environment", () => {
  process.env.OCA_TEST_IMG_RESOURCE = "myres";
  process.env.OCA_TEST_IMG_KEY = "azkey";
  try {
    const req = buildImageRequest(
      { type: "azure", baseUrl: "", model: "my image", apiKeyEnv: "OCA_TEST_IMG_KEY", resourceNameEnv: "OCA_TEST_IMG_RESOURCE" },
      "a cat",
      "landscape",
    );
    assert.equal(req.url, "https://myres.openai.azure.com/openai/deployments/my%20image/images/generations?api-version=2025-04-01-preview");
    assert.equal(req.headers["api-key"], "azkey");
    assert.equal(req.headers.authorization, undefined);
    assert.deepEqual(req.body, { prompt: "a cat", n: 1, size: "1536x1024" });
  } finally {
    delete process.env.OCA_TEST_IMG_RESOURCE;
    delete process.env.OCA_TEST_IMG_KEY;
  }
});

test("buildImageRequest: Gemini generateContent with the key header and an aspect ratio", () => {
  const req = buildImageRequest(
    { type: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", model: "gemini-3.1-flash-image", apiKey: "g-key" },
    "a cat",
    "landscape",
  );
  assert.equal(req.url, "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-image:generateContent");
  assert.equal(req.headers["x-goog-api-key"], "g-key");
  assert.equal(req.headers.authorization, undefined);
  assert.deepEqual(req.body, {
    contents: [{ parts: [{ text: "a cat" }] }],
    generationConfig: { responseModalities: ["IMAGE"], imageConfig: { aspectRatio: "3:2" } },
  });
});

test("geminiImageBytes takes the final image and explains an answer without one", () => {
  const part = (bytes: Buffer, thought?: boolean) => ({ inlineData: { mimeType: "image/png", data: bytes.toString("base64") }, thought });
  assert.deepEqual(geminiImageBytes({ candidates: [{ content: { parts: [{ text: "Here you go" }, part(JPG, true), part(PNG)] } }] }), PNG);
  assert.deepEqual(geminiImageBytes({ candidates: [{ content: { parts: [part(JPG, true)] } }] }), JPG);
  assert.throws(() => geminiImageBytes({ candidates: [{ finishReason: "IMAGE_SAFETY", content: { parts: [{ text: "I can't draw that." }] } }] }), /IMAGE_SAFETY.*can't draw/);
  assert.throws(() => geminiImageBytes({ promptFeedback: { blockReason: "PROHIBITED_CONTENT" } }), /PROHIBITED_CONTENT/);
});

test("generateImage reads a Gemini answer", async () => {
  const gemini: StoredImageProvider = { type: "gemini", baseUrl: "https://g.example/v1beta", model: "m", apiKey: "k" };
  const answer = Response.json({ candidates: [{ content: { parts: [{ inlineData: { mimeType: "image/jpeg", data: JPG.toString("base64") } }] } }] });
  const image = await generateImage(gemini, "a cat", "square", fakeFetch([answer]));
  assert.equal(image.ext, "jpg");
});

test("buildImageRequest rejects a provider without a base URL", () => {
  assert.throws(() => buildImageRequest({ type: "azure", baseUrl: "", model: "m", resourceNameEnv: "OCA_TEST_UNSET_VAR" }, "x", "square"), /OCA_TEST_UNSET_VAR/);
});

test("sniffImageExt recognizes common formats", () => {
  assert.equal(sniffImageExt(PNG), "png");
  assert.equal(sniffImageExt(JPG), "jpg");
  assert.equal(sniffImageExt(Buffer.from("RIFF\0\0\0\0WEBPVP8 ", "latin1")), "webp");
  assert.equal(sniffImageExt(Buffer.from("GIF89a", "latin1")), "gif");
});

test("generateImage decodes a base64 answer", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const image = await generateImage(openai, "a cat", "square", fakeFetch([Response.json({ data: [{ b64_json: PNG.toString("base64") }] })], calls));
  assert.deepEqual(image.bytes, PNG);
  assert.equal(image.ext, "png");
  assert.equal(calls.length, 1);
  assert.equal(calls[0].init?.method, "POST");
});

test("generateImage downloads a URL answer", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const image = await generateImage(openai, "a cat", "square", fakeFetch([Response.json({ data: [{ url: "https://img.example/x" }] }), new Response(JPG)], calls));
  assert.equal(image.ext, "jpg");
  assert.equal(calls[1].url, "https://img.example/x");
  // The download goes to another origin: the API key must not travel with it.
  assert.equal(calls[1].init?.headers, undefined);
});

test("generateImage surfaces the provider's error message", async () => {
  const failing = fakeFetch([Response.json({ error: { message: "Your request was rejected by the safety system." } }, { status: 400 })]);
  await assert.rejects(generateImage(openai, "x", "square", failing), /HTTP 400.*safety system/);
  await assert.rejects(generateImage(openai, "x", "square", fakeFetch([Response.json({ data: [] })])), /returned no image/);
});

test("pickOpenAiImageModels keeps image models, newest first, or everything when none match", () => {
  const data = [{ id: "gpt-6-luna", created: 9 }, { id: "gpt-image-2", created: 2 }, { id: "gpt-image-2.6-flare", created: 5 }, { id: "dall-e-3", created: 1 }];
  assert.deepEqual(pickOpenAiImageModels(data), ["gpt-image-2.6-flare", "gpt-image-2", "dall-e-3"]);
  assert.deepEqual(pickOpenAiImageModels([{ id: "sdxl" }, { id: "flux" }, {}]), ["flux", "sdxl"]);
});

test("pickGeminiImageModels keeps the image models generateContent can call", () => {
  const models = [
    { name: "models/gemini-3-pro", supportedGenerationMethods: ["generateContent"] },
    { name: "models/gemini-3.1-flash-image", supportedGenerationMethods: ["generateContent", "countTokens"] },
    { name: "models/gemini-nano-banana-2.1", supportedGenerationMethods: ["generateContent"] },
    { name: "models/imagen-4.0-generate-001", supportedGenerationMethods: ["predict"] },
  ];
  assert.deepEqual(pickGeminiImageModels(models), ["gemini-3.1-flash-image", "gemini-nano-banana-2.1"]);
});

test("listImageModels calls the provider's model list with its key", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const ids = await listImageModels(openai, fakeFetch([Response.json({ data: [{ id: "gpt-image-2.6-flare" }, { id: "gpt-6-luna" }] })], calls));
  assert.deepEqual(ids, ["gpt-image-2.6-flare"]);
  assert.equal(calls[0].url, "https://api.openai.com/v1/models");
  assert.equal((calls[0].init?.headers as Record<string, string>).authorization, "Bearer sk-test");

  const gemini: StoredImageProvider = { type: "gemini", baseUrl: "https://g.example/v1beta", model: "", apiKey: "k" };
  const answer = Response.json({ models: [{ name: "models/gemini-3.1-flash-image", supportedGenerationMethods: ["generateContent"] }] });
  assert.deepEqual(await listImageModels(gemini, fakeFetch([answer], calls)), ["gemini-3.1-flash-image"]);
  assert.equal(calls[1].url, "https://g.example/v1beta/models?pageSize=1000");
  assert.equal((calls[1].init?.headers as Record<string, string>)["x-goog-api-key"], "k");

  await assert.rejects(listImageModels(openai, fakeFetch([Response.json({ error: { message: "Incorrect API key" } }, { status: 401 })])), /HTTP 401: Incorrect API key/);
  await assert.rejects(listImageModels({ type: "azure", baseUrl: "https://r.openai.azure.com", model: "" }), /deployment name/);
});

test("imageProviderForListing reuses the stored key only for the stored endpoint", () => {
  const stored: StoredImageProvider = { type: "openai", baseUrl: "https://api.openai.com/v1", model: "m", apiKey: "stored", headers: { "x-a": "1" } };
  const same = imageProviderForListing(stored, { type: "openai", baseUrl: " https://api.openai.com/v1 " });
  assert.equal(same.apiKey, "stored");
  assert.deepEqual(same.headers, { "x-a": "1" });
  assert.equal(imageProviderForListing(stored, { type: "openai", baseUrl: "https://api.openai.com/v1", apiKey: " typed " }).apiKey, "typed");

  const elsewhere = imageProviderForListing(stored, { type: "openai", baseUrl: "https://evil.example/v1", apiKeyEnv: "OTHER_KEY" });
  assert.equal(elsewhere.apiKey, undefined);
  assert.equal(elsewhere.headers, undefined);
  assert.equal(elsewhere.apiKeyEnv, "OTHER_KEY");
  assert.equal(imageProviderForListing(stored, { type: "gemini", baseUrl: "https://api.openai.com/v1" }).apiKey, undefined);
});

test("saveGeneratedImage never overwrites", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "oca-img-"));
  try {
    const first = await saveGeneratedImage(dir, "cat", { bytes: PNG, ext: "png" });
    const second = await saveGeneratedImage(dir, "cat", { bytes: PNG, ext: "png" });
    assert.equal(path.basename(first), "cat.png");
    assert.equal(path.basename(second), "cat-2.png");
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("mergeImageProvider validates input and keeps the stored key unless one is given", () => {
  assert.throws(() => mergeImageProvider(undefined, { type: "openai", baseUrl: "https://x/v1", model: " " }), /model/);
  assert.throws(() => mergeImageProvider(undefined, { type: "openai", baseUrl: "x/v1", model: "m" }), /baseUrl/);
  assert.throws(() => mergeImageProvider(undefined, { type: "azure", baseUrl: "", model: "m" }), /baseUrl/);

  const first = mergeImageProvider(undefined, { type: "openai", baseUrl: " https://x/v1 ", model: " m ", apiKey: " k ", apiKeyEnv: "E" });
  assert.deepEqual(first, { type: "openai", baseUrl: "https://x/v1", model: "m", azureApiVersion: undefined, apiKey: "k", apiKeyEnv: "E", resourceNameEnv: undefined });

  const kept = mergeImageProvider({ ...first, headers: { "x-a": "1" } }, { type: "openai", baseUrl: "https://y/v1", model: "m2" });
  assert.equal(kept.apiKey, "k");
  assert.equal(kept.apiKeyEnv, "E");
  assert.deepEqual(kept.headers, { "x-a": "1" });

  const cleared = mergeImageProvider(first, { type: "openai", baseUrl: "https://x/v1", model: "m", apiKey: "", apiKeyEnv: "" });
  assert.equal(cleared.apiKey, undefined);
  assert.equal(cleared.apiKeyEnv, undefined);

  const azure = mergeImageProvider(undefined, { type: "azure", baseUrl: "", model: "dep", resourceNameEnv: "AZURE_RESOURCE_NAME" });
  assert.equal(azure.resourceNameEnv, "AZURE_RESOURCE_NAME");
});
