import {
  DEFAULT_PORT,
  PROTOCOL_VERSION,
  WS_PATH,
  type ClientMessage,
  type HostApp,
  type ServerMessage,
  type SessionSettings,
} from "@officecli-agent/protocol";
import { runOfficeOp } from "./office-ops";

export type ConnectionState = "disconnected" | "connecting" | "connected" | "ready";

export interface HostClientEvents {
  onState(state: ConnectionState): void;
  onMessage(message: ServerMessage): void;
}

export function defaultHostUrl(): string {
  return `wss://localhost:${DEFAULT_PORT}${WS_PATH}`;
}

/**
 * WebSocket connection to the local host process. Handles the hello handshake,
 * executes Office.js requests coming from the host, and reconnects with backoff.
 */
export class HostClient {
  private socket: WebSocket | undefined;
  private reconnectTimer: number | undefined;
  private attempts = 0;
  private stopped = false;

  constructor(
    private url: string,
    private readonly hostApp: HostApp,
    private readonly hello: () => Promise<{ documentName?: string; apiSets: Record<string, boolean>; settings: SessionSettings }>,
    private readonly events: HostClientEvents,
  ) {}

  connect(): void {
    this.stopped = false;
    this.open();
  }

  setUrl(url: string): void {
    this.url = url;
    this.attempts = 0;
    this.socket?.close();
  }

  stop(): void {
    this.stopped = true;
    if (this.reconnectTimer) window.clearTimeout(this.reconnectTimer);
    this.socket?.close();
  }

  send(message: ClientMessage): void {
    if (this.socket?.readyState === WebSocket.OPEN) this.socket.send(JSON.stringify(message));
  }

  private open(): void {
    this.events.onState("connecting");
    let socket: WebSocket;
    try {
      socket = new WebSocket(this.url);
    } catch (err) {
      this.scheduleReconnect();
      console.error(err);
      return;
    }
    this.socket = socket;

    socket.onopen = async () => {
      this.attempts = 0;
      this.events.onState("connected");
      const info = await this.hello();
      this.send({ type: "hello", protocolVersion: PROTOCOL_VERSION, hostApp: this.hostApp, ...info });
    };

    socket.onmessage = (ev) => {
      let message: ServerMessage;
      try {
        message = JSON.parse(String(ev.data)) as ServerMessage;
      } catch {
        return;
      }
      if (message.type === "office.request") {
        void this.handleOfficeRequest(message.requestId, message.op);
        return;
      }
      if (message.type === "ready") this.events.onState("ready");
      this.events.onMessage(message);
    };

    socket.onclose = () => {
      this.socket = undefined;
      this.events.onState("disconnected");
      this.scheduleReconnect();
    };
    socket.onerror = () => {
      /* onclose follows */
    };
  }

  private async handleOfficeRequest(requestId: string, op: Parameters<typeof runOfficeOp>[0]): Promise<void> {
    try {
      const result = await runOfficeOp(op, this.hostApp);
      this.send({ type: "office.result", requestId, ok: true, result });
    } catch (err) {
      this.send({ type: "office.result", requestId, ok: false, error: err instanceof Error ? err.message : String(err) });
    }
  }

  private scheduleReconnect(): void {
    if (this.stopped) return;
    const delay = Math.min(15_000, 1000 * 2 ** Math.min(this.attempts, 4));
    this.attempts += 1;
    this.reconnectTimer = window.setTimeout(() => this.open(), delay);
  }
}
