import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  COPILOT_PROVIDER,
  type AgentEvent,
  type ByokProviderInput,
  type HostApp,
  type ModelSummary,
  type PermissionSummary,
  type ProviderSummary,
  type ServerMessage,
  type SessionSettings,
} from "@officecli-agent/protocol";
import { ProviderEditor } from "./ProviderEditor";
import { HostClient, defaultHostUrl, isEmbeddedInWebView2, type ConnectionState } from "./host-client";
import { t } from "./i18n";

type ChatItem =
  | { id: string; role: "user"; text: string }
  | { id: string; role: "assistant"; text: string; streaming: boolean }
  | { id: string; role: "tool"; toolName: string; args?: unknown; status: "running" | "ok" | "error"; output?: string; mcpServer?: string }
  | { id: string; role: "system"; text: string };

interface Props {
  hostApp: HostApp;
}

const SETTINGS_KEY = "officecli-agent.settings";

function loadSettings(): { hostUrl: string; session: SessionSettings } {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (raw) return JSON.parse(raw) as { hostUrl: string; session: SessionSettings };
  } catch {
    /* ignore */
  }
  return { hostUrl: defaultHostUrl(), session: { provider: COPILOT_PROVIDER, autoApproveReadOnly: true } };
}

let counter = 0;
const nextId = () => `${Date.now()}-${counter++}`;

export function App({ hostApp: initialHostApp }: Props) {
  // The add-in knows the real application; the host echoes it back in `ready`.
  const [hostApp, setHostApp] = useState<HostApp>(initialHostApp);
  const [hostBuild, setHostBuild] = useState("");
  const [state, setState] = useState<ConnectionState>("disconnected");
  const [items, setItems] = useState<ChatItem[]>([]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [models, setModels] = useState<ModelSummary[]>([]);
  const [providers, setProviders] = useState<ProviderSummary[]>([]);
  const [editingProvider, setEditingProvider] = useState<ProviderSummary | "new" | null>(null);
  const [settings, setSettings] = useState(loadSettings);
  const [permissions, setPermissions] = useState<PermissionSummary[]>([]);
  const [status, setStatus] = useState<string>("");
  const [info, setInfo] = useState<{ workspaceDir: string; officecliAvailable: boolean; allowedDirectories: string[] } | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const listRef = useRef<HTMLDivElement>(null);
  const clientRef = useRef<HostClient | null>(null);
  const settingsRef = useRef(settings);
  settingsRef.current = settings;

  const append = useCallback((item: ChatItem) => setItems((prev) => [...prev, item]), []);

  const handleEvent = useCallback((event: AgentEvent) => {
    switch (event.kind) {
      case "turn.start":
        setBusy(true);
        break;
      case "message.delta":
        setItems((prev) => {
          const idx = prev.findIndex((i) => i.id === event.messageId);
          if (idx === -1) return [...prev, { id: event.messageId, role: "assistant", text: event.text, streaming: true }];
          const copy = [...prev];
          const cur = copy[idx] as Extract<ChatItem, { role: "assistant" }>;
          copy[idx] = { ...cur, text: cur.text + event.text };
          return copy;
        });
        break;
      case "message":
        setItems((prev) => {
          const idx = prev.findIndex((i) => i.id === event.messageId);
          const item: ChatItem = { id: event.messageId, role: "assistant", text: event.text, streaming: false };
          if (idx === -1) return event.text ? [...prev, item] : prev;
          const copy = [...prev];
          copy[idx] = item;
          return copy;
        });
        break;
      case "tool.start":
        append({ id: event.toolCallId, role: "tool", toolName: event.toolName, args: event.args, status: "running", mcpServer: event.mcpServer });
        break;
      case "tool.complete":
        setItems((prev) =>
          prev.map((i) =>
            i.id === event.toolCallId && i.role === "tool"
              ? { ...i, status: event.success ? "ok" : "error", output: event.error ?? event.output }
              : i,
          ),
        );
        break;
      case "idle":
        setBusy(false);
        break;
      case "error":
        setBusy(false);
        append({ id: nextId(), role: "system", text: t.error(event.message) });
        break;
      default:
        break;
    }
  }, [append]);

  const handleMessage = useCallback(
    (message: ServerMessage) => {
      switch (message.type) {
        case "ready":
          setHostApp(message.hostApp);
          setModels(message.models);
          setProviders(message.providers);
          setInfo({ workspaceDir: message.workspaceDir, officecliAvailable: message.officecliAvailable, allowedDirectories: message.allowedDirectories });
          setStatus(message.officecliAvailable ? "" : t.officecliMissing);
          setHostBuild(message.hostBuild ?? "");
          break;
        case "agent.event":
          handleEvent(message.event);
          break;
        case "permission.request":
          setPermissions((prev) => [...prev, message.permission]);
          break;
        case "providers":
          setProviders(message.providers);
          setEditingProvider(null);
          break;
        case "status":
          setStatus(message.message);
          break;
        case "error":
          setBusy(false);
          append({ id: nextId(), role: "system", text: t.error(message.message) });
          break;
        default:
          break;
      }
    },
    [append, handleEvent],
  );

  useEffect(() => {
    const client = new HostClient(settings.hostUrl, initialHostApp, () => settingsRef.current.session, { onState: setState, onMessage: handleMessage });
    clientRef.current = client;
    client.connect();
    return () => client.stop();
    // Reconnect only when the host URL changes; session settings travel with the next hello.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.hostUrl, initialHostApp]);

  useEffect(() => {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  }, [settings]);

  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight });
  }, [items, permissions]);

  const send = () => {
    const text = input.trim();
    if (!text || state !== "ready") return;
    append({ id: nextId(), role: "user", text });
    clientRef.current?.send({ type: "chat.send", text });
    setInput("");
    setBusy(true);
  };

  const abort = () => clientRef.current?.send({ type: "chat.abort" });

  const reset = () => {
    setItems([]);
    setBusy(false);
    setPermissions([]);
    // A new hello makes the host drop the old session and start one with the current settings.
    clientRef.current?.sendHello();
  };

  const answerPermission = (requestId: string, approved: boolean) => {
    clientRef.current?.send({ type: "permission.response", requestId, approved });
    setPermissions((prev) => prev.filter((p) => p.requestId !== requestId));
  };

  const updateSession = (patch: SessionSettings) => {
    const session = { ...settings.session, ...patch };
    setSettings({ ...settings, session });
    clientRef.current?.send({ type: "settings.update", settings: session });
  };

  const selectedProvider = providers.find((p) => p.name === (settings.session.provider ?? COPILOT_PROVIDER));
  const isCopilot = !selectedProvider || selectedProvider.type === "copilot";

  const stateLabel = useMemo(() => {
    switch (state) {
      case "ready":
        return t.stateReady;
      case "connected":
        return t.stateConnected;
      case "connecting":
        return t.stateConnecting;
      default:
        return isEmbeddedInWebView2() ? t.stateDisconnectedEmbedded : t.stateDisconnectedBrowser;
    }
  }, [state]);

  return (
    <div className="app">
      <header className="topbar">
        <div className="title">
          <span className={`dot ${state}`} />
          <strong>OfficeCLI Agent</strong>
          <span className="muted">{hostApp}</span>
        </div>
        <div className="actions">
          <button onClick={() => window.location.reload()} title={t.reloadTitle}>
            ↻
          </button>
          <button onClick={reset} title={t.resetTitle}>
            {t.reset}
          </button>
          <button onClick={() => setShowSettings((s) => !s)} title={showSettings ? t.closeSettings : t.openSettings} className={showSettings ? "active" : ""}>
            {t.settings}
          </button>
        </div>
      </header>

      {showSettings && (
        <section className="settings">
          {!clientRef.current?.embedded && (
            <label>
              {t.hostUrl}
              <input value={settings.hostUrl} onChange={(e) => setSettings({ ...settings, hostUrl: e.target.value })} />
            </label>
          )}
          <label>
            {t.provider}
            <select
              value={settings.session.provider ?? COPILOT_PROVIDER}
              onChange={(e) => updateSession({ provider: e.target.value, model: undefined })}
            >
              {providers.map((p) => (
                <option key={p.name} value={p.name}>
                  {p.type === "copilot" ? t.copilotProvider : `${p.name} — ${p.type}${p.hasApiKey ? "" : t.noApiKey}`}
                </option>
              ))}
            </select>
          </label>
          {isCopilot ? (
            <label>
              {t.model}
              <select value={settings.session.model ?? ""} onChange={(e) => updateSession({ model: e.target.value || undefined })}>
                <option value="">{t.defaultOption}</option>
                {models.map((m) => (
                  <option key={m.id} value={m.id}>
                    {m.name}
                  </option>
                ))}
              </select>
            </label>
          ) : (
            <label>
              {t.modelIdByok}
              <input
                value={settings.session.model ?? ""}
                placeholder={selectedProvider?.model ?? "gpt-6-luna / claude-sonnet-4.5 / llama3.1 …"}
                onChange={(e) => updateSession({ model: e.target.value || undefined })}
              />
            </label>
          )}
          <div className="provider-list">
            <div className="row-between">
              <span className="muted small">{t.byokProviders}</span>
              <button className="small-btn" onClick={() => setEditingProvider("new")}>
                {t.add}
              </button>
            </div>
            {providers
              .filter((p) => p.type !== "copilot")
              .map((p) => (
                <div key={p.name} className="provider-row">
                  <span>
                    <strong>{p.name}</strong> <span className="muted small">{p.type} · {p.baseUrl}</span>
                  </span>
                  <span>
                    <button className="small-btn" onClick={() => setEditingProvider(p)}>
                      {t.edit}
                    </button>
                    <button className="small-btn" onClick={() => clientRef.current?.send({ type: "provider.delete", name: p.name })}>
                      {t.delete}
                    </button>
                  </span>
                </div>
              ))}
            {editingProvider && (
              <ProviderEditor
                initial={editingProvider === "new" ? undefined : editingProvider}
                onCancel={() => setEditingProvider(null)}
                onSave={(input: ByokProviderInput) => clientRef.current?.send({ type: "provider.save", provider: input })}
              />
            )}
          </div>
          <label className="row">
            <input
              type="checkbox"
              checked={settings.session.autoApproveReadOnly ?? true}
              onChange={(e) => updateSession({ autoApproveReadOnly: e.target.checked })}
            />
            {t.autoApproveReadOnly}
          </label>
          {info && (
            <div className="info">
              <div>{t.workspaceDir}: <code>{info.workspaceDir}</code></div>
              <div>officecli: {info.officecliAvailable ? t.officecliAvailable : t.officecliNotFound}</div>
              <div>{t.hostBuild}: {hostBuild || t.unknown}</div>
              <div>{t.allowedDirs}: {info.allowedDirectories.length ? info.allowedDirectories.join(", ") : t.allowedDirsNone}</div>
            </div>
          )}
          <p className="muted small">
            {t.settingsNote}
            <code>~/.officecli-agent/config.json</code>
            {t.settingsNoteSuffix}
          </p>
          <div className="row-end">
            <button className="primary" onClick={() => setShowSettings(false)}>
              {t.close}
            </button>
          </div>
          <p className="muted small credits">
            Powered by{" "}
            <a href="https://github.com/iOfficeAI/OfficeCLI" target="_blank" rel="noreferrer">OfficeCLI</a> (Apache-2.0) and{" "}
            <a href="https://github.com/github/copilot-sdk" target="_blank" rel="noreferrer">GitHub Copilot SDK</a> (MIT).
          </p>
        </section>
      )}

      <div className="status">{status || stateLabel}</div>

      <main className="chat" ref={listRef}>
        {items.length === 0 && (
          <div className="empty">
            <p>{t.emptyIntro(hostApp)}</p>
            <ul>
              {t.examples.map((example) => (
                <li key={example}>{example}</li>
              ))}
            </ul>
          </div>
        )}
        {items.map((item) => (
          <ChatRow key={item.id} item={item} />
        ))}
        {busy && !items.some((i) => i.role === "assistant" && i.streaming) && (
          <div className="msg assistant thinking" aria-label={t.waiting}>
            <span className="dots">
              <span />
              <span />
              <span />
            </span>
          </div>
        )}
        {permissions.map((p) => (
          <div key={p.requestId} className="permission">
            <div className="permission-title">{p.title}</div>
            <pre>{p.detail}</pre>
            <div className="permission-actions">
              <button className="primary" onClick={() => answerPermission(p.requestId, true)}>
                {t.allow}
              </button>
              <button onClick={() => answerPermission(p.requestId, false)}>{t.deny}</button>
            </div>
          </div>
        ))}
      </main>

      <footer className="composer">
        <textarea
          value={input}
          placeholder={state === "ready" ? t.inputPlaceholder : t.inputDisabled}
          disabled={state !== "ready"}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          rows={3}
        />
        <div className="composer-actions">
          {busy ? (
            <button onClick={abort}>{t.stop}</button>
          ) : (
            <button className="primary" onClick={send} disabled={state !== "ready" || !input.trim()}>
              {t.send}
            </button>
          )}
        </div>
      </footer>
    </div>
  );
}

function ChatRow({ item }: { item: ChatItem }) {
  const [open, setOpen] = useState(false);
  switch (item.role) {
    case "user":
      return <div className="msg user">{item.text}</div>;
    case "assistant":
      return <div className={`msg assistant${item.streaming ? " streaming" : ""}`}>{item.text}</div>;
    case "system":
      return <div className="msg system">{item.text}</div>;
    case "tool": {
      const label = item.mcpServer && item.mcpServer !== item.toolName ? `${item.mcpServer}:${item.toolName}` : item.toolName;
      return (
        <div className={`tool ${item.status}`} onClick={() => setOpen((o) => !o)}>
          <span className="tool-icon">{item.status === "running" ? "⏳" : item.status === "ok" ? "✓" : "✗"}</span>
          <span className="tool-name">{label}</span>
          {open && (
            <pre className="tool-detail">
              {item.args !== undefined ? `args: ${JSON.stringify(item.args, null, 2)}\n` : ""}
              {item.output ? `\n${item.output}` : ""}
            </pre>
          )}
        </div>
      );
    }
  }
}
