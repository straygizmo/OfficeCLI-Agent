import {
  DEFAULT_PORT,
  PROTOCOL_VERSION,
  WS_PATH,
  type ClientMessage,
  type HostApp,
  type ServerMessage,
  type SessionSettings,
} from "@officecli-agent/protocol";
import { locale } from "./i18n";

export type ConnectionState = "disconnected" | "connecting" | "connected" | "ready";

export interface HostClientEvents {
  onState(state: ConnectionState): void;
  onMessage(message: ServerMessage): void;
}

// ---------------------------------------------------------------------------
// Transports
// ---------------------------------------------------------------------------

interface Transport {
  connect(): void;
  send(text: string): void;
  close(): void;
  /** True when the transport runs inside the VSTO add-in (WebView2). */
  readonly embedded: boolean;
}

interface TransportEvents {
  onOpen(): void;
  onMessage(text: string): void;
  onClose(): void;
}

interface WebView2Bridge {
  postMessage(message: unknown): void;
  addEventListener(type: "message", listener: (ev: { data: unknown }) => void): void;
  removeEventListener(type: "message", listener: (ev: { data: unknown }) => void): void;
}

declare global {
  interface Window {
    chrome?: { webview?: WebView2Bridge };
  }
}

export function isEmbeddedInWebView2(): boolean {
  return typeof window !== "undefined" && Boolean(window.chrome?.webview);
}

/**
 * Inside the VSTO task pane the page talks to the C# add-in through
 * window.chrome.webview. The add-in relays frames to/from the host process and
 * executes Office operations itself, so the page never touches the document.
 */
class WebView2Transport implements Transport {
  readonly embedded = true;
  private listener?: (ev: { data: unknown }) => void;

  constructor(private readonly events: TransportEvents) {}

  connect(): void {
    const bridge = window.chrome!.webview!;
    this.listener = (ev) => {
      const text = typeof ev.data === "string" ? ev.data : JSON.stringify(ev.data);
      // The add-in reports relay state with a tiny control frame.
      if (text.startsWith('{"__bridge"')) {
        const ctl = JSON.parse(text) as { __bridge: "open" | "closed" };
        if (ctl.__bridge === "open") this.events.onOpen();
        else this.events.onClose();
        return;
      }
      this.events.onMessage(text);
    };
    bridge.addEventListener("message", this.listener);
    bridge.postMessage({ __bridge: "ready" });
  }

  send(text: string): void {
    window.chrome!.webview!.postMessage(text);
  }

  close(): void {
    if (this.listener) window.chrome!.webview!.removeEventListener("message", this.listener);
  }
}

/** Plain WebSocket, used when the UI is opened in a browser for debugging. */
class WebSocketTransport implements Transport {
  readonly embedded = false;
  private socket?: WebSocket;
  private timer?: number;
  private attempts = 0;
  private stopped = false;

  constructor(
    private url: string,
    private readonly events: TransportEvents,
  ) {}

  connect(): void {
    this.stopped = false;
    this.open();
  }

  send(text: string): void {
    if (this.socket?.readyState === WebSocket.OPEN) this.socket.send(text);
  }

  close(): void {
    this.stopped = true;
    if (this.timer) window.clearTimeout(this.timer);
    this.socket?.close();
  }

  private open(): void {
    let socket: WebSocket;
    try {
      socket = new WebSocket(this.url);
    } catch {
      this.scheduleReconnect();
      return;
    }
    this.socket = socket;
    socket.onopen = () => {
      this.attempts = 0;
      this.events.onOpen();
    };
    socket.onmessage = (ev) => this.events.onMessage(String(ev.data));
    socket.onclose = () => {
      this.socket = undefined;
      this.events.onClose();
      this.scheduleReconnect();
    };
  }

  private scheduleReconnect(): void {
    if (this.stopped) return;
    const delay = Math.min(15_000, 1000 * 2 ** Math.min(this.attempts, 4));
    this.attempts += 1;
    this.timer = window.setTimeout(() => this.open(), delay);
  }
}

// ---------------------------------------------------------------------------

export function defaultHostUrl(): string {
  return `ws://127.0.0.1:${DEFAULT_PORT}${WS_PATH}`;
}

/** Host app for browser debugging (?hostApp=Excel). The add-in overrides this in `hello`. */
export function hostAppFromQuery(): HostApp {
  const v = new URLSearchParams(window.location.search).get("hostApp");
  return v === "Excel" || v === "PowerPoint" ? v : "Word";
}

/**
 * Connection to the host. Sends the hello handshake on open; inside the add-in
 * the C# relay fills in the real host application and document name.
 */
export class HostClient {
  private transport: Transport;
  readonly embedded: boolean;

  constructor(
    url: string,
    private readonly hostApp: HostApp,
    private readonly settings: () => SessionSettings,
    private readonly events: HostClientEvents,
  ) {
    const transportEvents: TransportEvents = {
      onOpen: () => {
        this.events.onState("connected");
        this.sendHello();
      },
      onMessage: (text) => {
        let message: ServerMessage;
        try {
          message = JSON.parse(text) as ServerMessage;
        } catch {
          return;
        }
        if (message.type === "ready") this.events.onState("ready");
        this.events.onMessage(message);
      },
      onClose: () => this.events.onState("disconnected"),
    };
    this.embedded = isEmbeddedInWebView2();
    this.transport = this.embedded ? new WebView2Transport(transportEvents) : new WebSocketTransport(url, transportEvents);
  }

  connect(): void {
    this.events.onState("connecting");
    this.transport.connect();
  }

  stop(): void {
    this.transport.close();
  }

  send(message: ClientMessage): void {
    this.transport.send(JSON.stringify(message));
  }

  /** Ask the add-in to open a file or folder in Windows. Returns false outside the task pane (nothing happens). */
  openPath(path: string): boolean {
    if (!this.embedded) return false;
    window.chrome!.webview!.postMessage({ __bridge: "openPath", path });
    return true;
  }

  /** Start a fresh session (the host tears down the old one on a new hello). */
  sendHello(): void {
    this.send({ type: "hello", protocolVersion: PROTOCOL_VERSION, hostApp: this.hostApp, locale, settings: this.settings() });
  }
}
