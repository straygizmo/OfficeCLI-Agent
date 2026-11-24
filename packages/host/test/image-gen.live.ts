/**
 * Manual check against the real providers (costs a little; not part of `pnpm test`):
 *   node --import tsx test/image-gen.live.ts <outDir> [openai|gemini]...
 * Keys come from OPENAI_API_KEY / GEMINI_API_KEY.
 */
import type { StoredImageProvider } from "../src/config.js";
import { generateImage, saveGeneratedImage, type ImageAspect } from "../src/image-gen.js";

const PROVIDERS: Record<string, StoredImageProvider> = {
  openai: { type: "openai", baseUrl: "https://api.openai.com/v1", model: process.env.OPENAI_IMAGE_MODEL ?? "gpt-image-2.6-flare", apiKeyEnv: "OPENAI_API_KEY" },
  gemini: { type: "gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta", model: process.env.GEMINI_IMAGE_MODEL ?? "gemini-3.1-flash-image", apiKeyEnv: "GEMINI_API_KEY" },
};

const [outDir, ...names] = process.argv.slice(2);
if (!outDir) throw new Error("usage: image-gen.live.ts <outDir> [openai|gemini]...");

let failed = false;
for (const name of names.length ? names : Object.keys(PROVIDERS)) {
  for (const aspect of ["square", "landscape"] as ImageAspect[]) {
    const started = Date.now();
    try {
      const image = await generateImage(PROVIDERS[name], "A simple flat illustration of a red lighthouse on a cliff at sunset, no text", aspect);
      const file = await saveGeneratedImage(outDir, `${name}-${aspect}`, image);
      console.log(`${name} ${aspect}: ok ${file} (${image.bytes.length} bytes, ${Math.round((Date.now() - started) / 1000)}s)`);
    } catch (err) {
      failed = true;
      console.log(`${name} ${aspect}: FAILED ${(err as Error).message}`);
    }
  }
}
process.exitCode = failed ? 1 : 0;
