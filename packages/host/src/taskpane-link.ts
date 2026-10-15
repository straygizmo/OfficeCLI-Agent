import { randomUUID } from "node:crypto";
import type { WebSocket } from "ws";
import type { ClientMessage, OfficeOp, PermissionSummary, ServerMessage } from "@officecli-agent/protocol";

interface Pending<T> {
  resolve: (value: T) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

/**
 * One connected task pane. Owns the socket, correlates request/response pairs
 * (Office.js operations, permission prompts) and exposes them as promises.
 */
export class TaskpaneLink {
  private officePending = new Map<string, Pending<unknown>>();
  private permissionPending = new Map<string, Pending<boolean>>();
  private closed = false;

  constructor(
    private readonly socket: WebSocket,
    private readonly officeTimeoutMs = 120_000,
  ) {}

  send(message: ServerMessage): void {
    if (this.closed || this.socket.readyState !== this.socket.OPEN) return;
    this.socket.send(JSON.stringify(message));
  }

  status(message: string): void {
    this.send({ type: "status", message });
  }

  /** Ask the task pane to run an Office.js operation and wait for the result. */
  runOffice<T = unknown>(op: OfficeOp): Promise<T> {
    const requestId = randomUUID();
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.officePending.delete(requestId);
        reject(new Error(`Office operation '${op.kind}' timed out after ${this.officeTimeoutMs / 1000}s`));
      }, this.officeTimeoutMs);
      this.officePending.set(requestId, { resolve: resolve as (v: unknown) => void, reject, timer });
      this.send({ type: "office.request", requestId, op });
    });
  }

  /** Show a permission prompt in the task pane; resolves false on timeout. */
  askPermission(summary: Omit<PermissionSummary, "requestId">, timeoutMs: number): Promise<boolean> {
    const requestId = randomUUID();
    return new Promise<boolean>((resolve) => {
      const timer = setTimeout(() => {
        this.permissionPending.delete(requestId);
        this.status("Permission request timed out; denied.");
        resolve(false);
      }, timeoutMs);
      this.permissionPending.set(requestId, { resolve, reject: () => resolve(false), timer });
      this.send({ type: "permission.request", permission: { requestId, ...summary } });
    });
  }

  /** Feed an inbound message; returns true when it was a response we were waiting for. */
  handleResponse(message: ClientMessage): boolean {
    if (message.type === "office.result") {
      const pending = this.officePending.get(message.requestId);
      if (!pending) return true;
      this.officePending.delete(message.requestId);
      clearTimeout(pending.timer);
      if (message.ok) pending.resolve(message.result);
      else pending.reject(new Error(message.error));
      return true;
    }
    if (message.type === "permission.response") {
      const pending = this.permissionPending.get(message.requestId);
      if (!pending) return true;
      this.permissionPending.delete(message.requestId);
      clearTimeout(pending.timer);
      pending.resolve(message.approved);
      return true;
    }
    return false;
  }

  close(): void {
    this.closed = true;
    for (const p of this.officePending.values()) {
      clearTimeout(p.timer);
      p.reject(new Error("Task pane disconnected"));
    }
    for (const p of this.permissionPending.values()) {
      clearTimeout(p.timer);
      p.resolve(false);
    }
    this.officePending.clear();
    this.permissionPending.clear();
  }
}
