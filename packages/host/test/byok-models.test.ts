import { test } from "node:test";
import assert from "node:assert/strict";
import { firstServerModel } from "../src/byok-models.js";

function fakeFetch(response: Response, calls: Array<{ url: string; init?: RequestInit }> = []): typeof fetch {
  return (async (url: string | URL | Request, init?: RequestInit) => {
    calls.push({ url: String(url), init });
    return response;
  }) as typeof fetch;
}

test("firstServerModel takes the first listed model id", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  const res = new Response(JSON.stringify({ data: [{ id: "" }, { id: "llama3.1:8b" }, { id: "qwen3" }] }), { status: 200 });
  assert.equal(await firstServerModel("http://localhost:11434/v1/", undefined, undefined, fakeFetch(res, calls)), "llama3.1:8b");
  assert.equal(calls[0].url, "http://localhost:11434/v1/models");
  assert.equal((calls[0].init?.headers as Record<string, string>).authorization, undefined);
});

test("firstServerModel sends the key and reports failures", async () => {
  const calls: Array<{ url: string; init?: RequestInit }> = [];
  await assert.rejects(firstServerModel("http://x/v1", "k", { "x-a": "1" }, fakeFetch(new Response("no", { status: 401 }), calls)), /HTTP 401/);
  assert.deepEqual(calls[0].init?.headers, { "x-a": "1", authorization: "Bearer k" });
  await assert.rejects(firstServerModel("http://x/v1", undefined, undefined, fakeFetch(new Response(JSON.stringify({ data: [] })))), /no models/);
});
