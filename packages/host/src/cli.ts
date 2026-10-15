#!/usr/bin/env node
import { CONFIG_FILE, loadConfig, parseArgs } from "./config.js";
import { startHost } from "./server.js";

async function main(): Promise<void> {
  const config = loadConfig(parseArgs(process.argv.slice(2)));
  console.log(`[host] config: ${CONFIG_FILE}`);
  const server = await startHost(config);

  const shutdown = async (signal: string) => {
    console.log(`\n[host] ${signal} received, shutting down…`);
    // The Copilot runtime can take a while to stop; never hang the process on it.
    setTimeout(() => process.exit(0), 5000).unref();
    try {
      await server.close();
    } finally {
      process.exit(0);
    }
  };
  process.on("SIGINT", () => void shutdown("SIGINT"));
  process.on("SIGTERM", () => void shutdown("SIGTERM"));
}

main().catch((err) => {
  console.error("[host] fatal:", err);
  process.exit(1);
});
