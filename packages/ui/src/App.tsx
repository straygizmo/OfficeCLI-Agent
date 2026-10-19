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
  // Settings the current session was started with (edits apply only after a reset).
  const [activeSession, setActiveSession] = useState<SessionSettings>(settings.session);
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
    const client = new HostClient(settings.hostUrl, initialHostApp, () => {
      const session = settingsRef.current.session;
      setActiveSession(session);
      return session;
    }, { onState: setState, onMessage: handleMessage });
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

  const activeLabel = useMemo(() => {
    const provider = providers.find((p) => p.name === (activeSession.provider ?? COPILOT_PROVIDER));
    const copilot = !provider || provider.type === "copilot";
    const providerName = copilot ? "Copilot" : provider.name;
    const modelId = activeSession.model ?? (copilot ? undefined : provider.model);
    const modelName = copilot ? models.find((m) => m.id === modelId)?.name ?? modelId : modelId;
    return `${providerName} / ${modelName ?? t.defaultModelShort}`;
  }, [providers, models, activeSession]);

  const [stateLabel, stateTip] = useMemo(() => {
    switch (state) {
      case "ready":
        return [t.stateReady, t.stateReadyTip];
      case "connected":
        return [t.stateConnected, t.stateConnectedTip];
      case "connecting":
        return [t.stateConnecting, t.stateConnectingTip];
      default:
        return [t.stateDisconnected, isEmbeddedInWebView2() ? t.stateDisconnectedEmbeddedTip : t.stateDisconnectedBrowserTip];
    }
  }, [state]);

  return (
    <div className="app">
      <header className="topbar">
        <span className="active-model muted small" title={activeLabel}>
          {activeLabel}
        </span>
        <div className="actions">
          <span className={`badge ${state}`} title={stateTip}>
            {stateLabel}
          </span>
          <button className="icon" onClick={() => window.location.reload()} title={t.reloadTitle} aria-label={t.reloadTitle}>
            ↻
          </button>
          <button className="icon" onClick={reset} title={t.resetTitle} aria-label={t.reset}>
            <svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M2.5 4h11M6 4V2.5h4V4M4 4l.7 9.5h6.6L12 4M6.8 6.5v4.5M9.2 6.5v4.5" />
            </svg>
          </button>
          <button
            onClick={() => setShowSettings((s) => !s)}
            title={showSettings ? t.closeSettings : t.openSettings}
            aria-label={t.settings}
            className={`icon${showSettings ? " active" : ""}`}
          >
            <svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <circle cx="8" cy="8" r="2.2" />
              <path d="M8 1.5v1.8M8 12.7v1.8M1.5 8h1.8M12.7 8h1.8M3.4 3.4l1.3 1.3M11.3 11.3l1.3 1.3M3.4 12.6l1.3-1.3M11.3 4.7l1.3-1.3" />
              <circle cx="8" cy="8" r="4.3" />
            </svg>
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

      {status && <div className="status">{status}</div>}

      <main className="chat" ref={listRef}>
        {items.length === 0 && (
          <div className="empty">
            <p>{t.emptyIntro[hostApp]}</p>
            <ul>
              {t.examples[hostApp].map((example) => (
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
        <div className="input-wrap">
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
            rows={4}
          />
          <div className="input-action">
            {busy ? (
              <button className="icon stop" onClick={abort} title={t.stop} aria-label={t.stop}>
                <svg viewBox="0 0 16 16" width="14" height="14" fill="currentColor" aria-hidden="true">
                  <rect x="3.5" y="3.5" width="9" height="9" rx="1.5" />
                </svg>
              </button>
            ) : (
              <button className="icon" onClick={send} disabled={state !== "ready" || !input.trim()} title={t.send} aria-label={t.send}>
                <svg viewBox="0 0 16 16" width="15" height="15" fill="currentColor" aria-hidden="true">
                  <path d="M1.7 2.2a.6.6 0 0 1 .8-.7l11.8 5.9a.7.7 0 0 1 0 1.2L2.5 14.5a.6.6 0 0 1-.8-.7L3 8.6h5.2a.6.6 0 0 0 0-1.2H3L1.7 2.2z" />
                </svg>
              </button>
            )}
          </div>
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
