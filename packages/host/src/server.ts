import { execFile } from "node:child_process";
import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { CopilotClient, RuntimeConnection, type ModelInfo } from "@github/copilot-sdk";
import { WebSocketServer, type WebSocket } from "ws";
import { ATTACHMENT_MAX_BYTES, ATTACHMENT_MAX_FILES, PROTOCOL_VERSION, WS_PATH, isClientMessage, type ClientMessage, type ModelSummary } from "@officecli-agent/protocol";
import { AgentSession } from "./agent-session.js";
import { deleteProvider, setMcpServerEnabled, setSkillEnabled, summarizeProviders, upsertProvider, type HostConfig } from "./config.js";
import { TaskpaneLink } from "./taskpane-link.js";
import { applyToggleToSessions } from "./toggles.js";

const execFileAsync = promisify(execFile);

const MIME: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".map": "application/json",
};

export interface HostServer {
  close(): Promise<void>;
}

/** Build stamp (mtime of this file) so a stale host process is easy to spot in logs and the UI. */
export function hostBuildStamp(): string {
  try {
    return fs.statSync(fileURLToPath(import.meta.url)).mtime.toISOString().replace("T", " ").slice(0, 19);
  } catch {
    return "unknown";
  }
}

export async function startHost(config: HostConfig): Promise<HostServer> {
  console.log(`[host] build ${hostBuildStamp()} (pid ${process.pid})`);
  const officecliAvailable = await checkOfficecli(config.officecliPath);
  if (!officecliAvailable) {
    console.warn(`[host] officecli not found at '${config.officecliPath}'. Advanced document tools will be disabled.`);
  }

  const client = new CopilotClient({
    connection: RuntimeConnection.forStdio(config.copilotCliPath ? { path: config.copilotCliPath } : undefined),
    workingDirectory: config.workspaceRoot,
    logLevel: "warning",
  });
  await client.start();
  console.log("[host] Copilot CLI runtime started");

  let models: ModelSummary[] = [];
  try {
    const list: ModelInfo[] = await client.listModels();
    models = list.map((m) => ({ id: m.id, name: m.name }));
  } catch (err) {
    console.warn(`[host] listModels failed: ${(err as Error).message}`);
  }

  const taskpaneDir = resolveTaskpaneDir();
  const httpServer = createHttpServer(config, (req, res) => serveStatic(req, res, taskpaneDir, config.serveTaskpane));
  // Attachments arrive as base64 inside chat.send frames; allow the full batch plus slack.
  const wss = new WebSocketServer({ server: httpServer, path: WS_PATH, maxPayload: Math.ceil((ATTACHMENT_MAX_FILES * ATTACHMENT_MAX_BYTES * 4) / 3) + 1024 * 1024 });
  const sessions = new Set<AgentSession>();

  /** After a toggle, every connected pane gets the fresh catalog (the choice is shared through config.json). */
  const broadcastCatalog = async () => {
    for (const s of sessions) {
      try {
        s.link.send({ type: "catalog", catalog: await s.refreshCatalog() });
      } catch (err) {
        console.warn(`[host] catalog refresh failed: ${(err as Error).message}`);
      }
    }
  };

  // Idle shutdown: the VSTO add-in starts us with --idle-exit so we go away once
  // every Office window that used us has closed.
  let idleTimer: NodeJS.Timeout | undefined;
  const armIdleTimer = () => {
    if (!config.idleExitSeconds || wss.clients.size > 0) return;
    idleTimer = setTimeout(() => {
      if (wss.clients.size > 0) return;
      console.log(`[host] idle for ${config.idleExitSeconds}s, exiting`);
      process.emit("SIGTERM");
    }, config.idleExitSeconds * 1000);
  };
  armIdleTimer();

  wss.on("connection", (socket: WebSocket) => {
    if (idleTimer) {
      clearTimeout(idleTimer);
      idleTimer = undefined;
    }
    const link = new TaskpaneLink(socket);
    let session: AgentSession | undefined;
    console.log("[host] add-in connected");

    socket.on("message", async (raw) => {
      let message: unknown;
      try {
        message = JSON.parse(raw.toString());
      } catch {
        link.send({ type: "error", message: "Invalid JSON" });
        return;
      }
      if (!isClientMessage(message)) return;
      if (link.handleResponse(message)) return;

      try {
        session = await handleMessage(message, session);
      } catch (err) {
        const text = (err as Error).message ?? String(err);
        console.error(`[host] ${text}`);
        link.send({ type: "error", message: text });
      }
    });

    async function handleMessage(message: ClientMessage, current: AgentSession | undefined): Promise<AgentSession | undefined> {
      switch (message.type) {
        case "hello": {
          if (message.protocolVersion !== PROTOCOL_VERSION) {
            link.send({ type: "error", message: `Protocol mismatch: task pane v${message.protocolVersion}, host v${PROTOCOL_VERSION}. Rebuild the add-in.` });
          }
          if (current) {
            sessions.delete(current);
            await current.destroy();
          }
          link.status(`Starting Copilot session… (host build ${hostBuildStamp()})`);
          const next = new AgentSession({
            client,
            config,
            link,
            hostApp: message.hostApp,
            documentName: message.documentName,
            locale: message.locale ?? "en",
            settings: message.settings ?? {},
            officecliAvailable,
          });
          await next.start();
          sessions.add(next);
          link.send({
            type: "ready",
            sessionId: next.sessionId,
            hostApp: message.hostApp,
            hostBuild: hostBuildStamp(),
            models,
            providers: summarizeProviders(config),
            settings: {
              provider: message.settings?.provider ?? config.defaultProvider,
              model: message.settings?.model ?? config.model,
              autoApproveReadOnly: message.settings?.autoApproveReadOnly ?? config.autoApproveReadOnly,
              autoApproveOfficecli: message.settings?.autoApproveOfficecli ?? false,
              workAcrossOpenFiles: message.settings?.workAcrossOpenFiles ?? false,
              webSearch: message.settings?.webSearch ?? true,
            },
            workspaceDir: next.workspaceDir,
            allowedDirectories: config.allowedDirectories,
            officecliAvailable,
            catalog: next.catalog,
          });
          return next;
        }
        case "chat.send":
          if (!current) throw new Error("Session not ready; send 'hello' first.");
          await current.send(message.text, message.attachments);
          return current;
        case "chat.abort":
          await current?.abort();
          return current;
        case "settings.update":
          // Session-creation options (provider, model, ...) wait for the next hello; the task pane says when.
          current?.updateSettings(message.settings);
          return current;
        case "provider.save":
          upsertProvider(config, message.provider);
          link.send({ type: "providers", providers: summarizeProviders(config) });
          link.status(`Provider '${message.provider.name}' saved.`);
          return current;
        case "provider.delete":
          deleteProvider(config, message.name);
          link.send({ type: "providers", providers: summarizeProviders(config) });
          link.status(`Provider '${message.name}' removed.`);
          return current;
        case "mcp.setEnabled":
          setMcpServerEnabled(config, message.name, message.enabled);
          await applyToggleToSessions(sessions, "mcp", message.name, message.enabled);
          await broadcastCatalog();
          return current;
        case "skills.setEnabled":
          setSkillEnabled(config, message.name, message.enabled);
          await applyToggleToSessions(sessions, "skill", message.name, message.enabled);
          await broadcastCatalog();
          return current;
        case "session.reset":
          if (current) {
            sessions.delete(current);
            await current.destroy();
          }
          link.status("Session cleared. Send a message to start a new one.");
          return undefined;
        default:
          return current;
      }
    }

    socket.on("close", async () => {
      console.log("[host] add-in disconnected");
      link.close();
      if (session) {
        sessions.delete(session);
        await session.destroy();
      }
      armIdleTimer();
    });
  });

  await new Promise<void>((resolve) => httpServer.listen(config.port, "127.0.0.1", resolve));
  const scheme = httpServer instanceof https.Server ? "https" : "http";
  console.log(`[host] listening on ${scheme}://localhost:${config.port}  (ws path ${WS_PATH})`);
  if (config.serveTaskpane) console.log(`[host] chat UI (debug): ${scheme}://localhost:${config.port}/taskpane.html?hostApp=Word`);
  console.log("[host] READY");

  return {
    async close() {
      for (const s of sessions) await s.destroy();
      wss.close();
      await new Promise<void>((resolve) => httpServer.close(() => resolve()));
      await client.stop();
    },
  };
}

// ---------------------------------------------------------------------------

async function checkOfficecli(bin: string): Promise<boolean> {
  try {
    const { stdout } = await execFileAsync(bin, ["--version"], { timeout: 15_000, windowsHide: true });
    console.log(`[host] officecli ${stdout.trim()}`);
    return true;
  } catch {
    return false;
  }
}

function createHttpServer(config: HostConfig, handler: http.RequestListener): http.Server | https.Server {
  if (!config.tls) return http.createServer(handler);
  const cert = path.join(config.certDir, "localhost.crt");
  const key = path.join(config.certDir, "localhost.key");
  if (fs.existsSync(cert) && fs.existsSync(key)) {
    return https.createServer({ cert: fs.readFileSync(cert), key: fs.readFileSync(key) }, handler);
  }
  console.warn(`[host] tls=true but no certificate in ${config.certDir}; falling back to plain HTTP.`);
  return http.createServer(handler);
}

function resolveTaskpaneDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // packages/host/dist -> packages/ui/dist
  return path.resolve(here, "..", "..", "ui", "dist");
}

function serveStatic(req: http.IncomingMessage, res: http.ServerResponse, root: string, enabled: boolean): void {
  if (!enabled) {
    res.writeHead(404).end();
    return;
  }
  const url = new URL(req.url ?? "/", "http://localhost");
  let rel = decodeURIComponent(url.pathname);
  if (rel === "/" || rel === "") rel = "/taskpane.html";
  const file = path.normalize(path.join(root, rel));
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.writeHead(404, { "content-type": "text/plain" }).end(`Not found: ${rel}\nBuild the add-in first (pnpm build).`);
    return;
  }
  res.writeHead(200, {
    "content-type": MIME[path.extname(file).toLowerCase()] ?? "application/octet-stream",
    "cache-control": "no-cache",
  });
  fs.createReadStream(file).pipe(res);
}
