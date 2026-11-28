import { test } from "node:test";
import assert from "node:assert/strict";
import { describeAuthError, errorReason, isAuthError, modelsRequest, probeAuthFailure, type ProbeTarget } from "../src/provider-errors.js";

const openai: ProbeTarget = { name: "OpenAI", type: "openai", baseUrl: "https://api.openai.com/v1/", apiKey: "sk-x" };
const runtimeMessage =
  "Authentication failed with provider at https://api.openai.com/v1 (HTTP 401).\n  Check your COPILOT_PROVIDER_API_KEY, COPILOT_PROVIDER_API_KEY_COMMAND, or COPILOT_PROVIDER_BEARER_TOKEN.";

test("isAuthError recognizes the runtime's auth failures only", () => {
  assert.equal(isAuthError({ message: runtimeMessage }), true);
  assert.equal(isAuthError({ message: "x", errorType: "authentication" }), true);
  assert.equal(isAuthError({ message: "x", statusCode: 403 }), true);
  assert.equal(isAuthError({ message: "rate limited", errorType: "rate_limit", statusCode: 429 }), false);
});

test("modelsRequest authenticates the way each provider type expects", () => {
  assert.deepEqual(modelsRequest(openai), { url: "https://api.openai.com/v1/models", headers: { authorization: "Bearer sk-x" } });
  const anthropic = modelsRequest({ name: "a", type: "anthropic", baseUrl: "https://api.anthropic.com", apiKey: "k" });
  assert.equal(anthropic.url, "https://api.anthropic.com/v1/models");
  assert.equal(anthropic.headers["x-api-key"], "k");
  for (const baseUrl of ["https://r.openai.azure.com", "https://r.openai.azure.com/openai/v1/"]) {
    const azure = modelsRequest({ name: "z", type: "azure", baseUrl, apiKey: "k" });
    assert.equal(azure.url, "https://r.openai.azure.com/openai/v1/models");
    assert.equal(azure.headers["api-key"], "k");
  }
});

test("errorReason reads OpenAI / Anthropic style bodies", () => {
  assert.equal(
    errorReason(JSON.stringify({ error: { message: "Your API key has expired. Create a new API key to continue.", code: "expired_secret_key" } })),
    "Your API key has expired. Create a new API key to continue. (expired_secret_key)",
  );
  assert.equal(errorReason(JSON.stringify({ type: "error", error: { type: "authentication_error", message: "invalid x-api-key" } })), "invalid x-api-key");
  assert.equal(errorReason("<html>nope</html>"), undefined);
});

test("probeAuthFailure returns the provider's status and reason, or nothing", async () => {
  const body = JSON.stringify({ error: { message: "Your API key has expired.", code: "expired_secret_key" } });
  const fail = (async () => new Response(body, { status: 401 })) as typeof fetch;
  assert.deepEqual(await probeAuthFailure(openai, fail), { status: 401, reason: "Your API key has expired. (expired_secret_key)" });
  const ok = (async () => new Response("{}", { status: 200 })) as typeof fetch;
  assert.equal(await probeAuthFailure(openai, ok), undefined);
  const down = (async () => {
    throw new Error("ECONNREFUSED");
  }) as typeof fetch;
  assert.equal(await probeAuthFailure(openai, down), undefined);
});

test("describeAuthError names the provider, the reason and the fix, not the runtime's env vars", () => {
  const ja = describeAuthError("ja", openai, { message: runtimeMessage, statusCode: 401 }, { status: 401, reason: "Your API key has expired." });
  assert.match(ja, /「OpenAI」が API キーを受け付けませんでした \(HTTP 401\)/);
  assert.match(ja, /理由: Your API key has expired\./);
  assert.match(ja, /BYOK プロバイダー/);
  assert.doesNotMatch(ja, /COPILOT_PROVIDER/);
  const en = describeAuthError("en", { ...openai, apiKey: undefined }, { message: runtimeMessage, providerErrorCode: "missing_key" }, undefined);
  assert.match(en, /has no API key/);
  assert.match(en, /Reason: missing_key/);
});
