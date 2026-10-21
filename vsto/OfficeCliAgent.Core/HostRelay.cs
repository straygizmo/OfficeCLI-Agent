using System;
using System.IO;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace OfficeCliAgent.Core
{
    /// <summary>What a frame posted by the page is, judged from its first bytes only.</summary>
    public enum PageFrameKind { Bridge, Hello, Other }

    /// <summary>
    /// Bridges the chat page (WebView2) and the Node host (WebSocket).
    ///
    /// Frames flow through unchanged in both directions, with three exceptions:
    ///  * the page's <c>hello</c> is stamped with the real host application and document;
    ///  * <c>{"__bridge":…}</c> control frames from the page are handled here (relay state, open a path);
    ///  * <c>office.request</c> frames from the host are executed here through COM
    ///    (on the Office UI thread) and answered with <c>office.result</c>.
    /// Large chat frames (attachments) are never parsed here.
    /// </summary>
    public sealed class HostRelay : IDisposable
    {
        private readonly AgentSettings _settings;
        private readonly IOfficeOps _ops;
        private readonly Control _uiThread;
        private readonly Action<string> _postToPage;
        private readonly Func<CancellationToken, Task> _ensureHost;
        private readonly SemaphoreSlim _sendGate = new SemaphoreSlim(1, 1);
        private ClientWebSocket _socket;
        private volatile bool _connected;

        /// <param name="ensureHost">Called before every connection attempt; starts the host process when nothing listens.</param>
        public HostRelay(AgentSettings settings, IOfficeOps ops, Control uiThread, Action<string> postToPage, Func<CancellationToken, Task> ensureHost)
        {
            _settings = settings;
            _ops = ops;
            _uiThread = uiThread;
            _postToPage = postToPage;
            _ensureHost = ensureHost;
        }

        public bool IsConnected => _connected;

        /// <summary>Connect and keep reconnecting until cancelled.</summary>
        public async Task RunAsync(CancellationToken ct)
        {
            var uri = new Uri($"ws://127.0.0.1:{_settings.Port}/ws");
            var failures = 0;
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    // (Re)start the host if it is gone, e.g. after an idle exit or a manual kill during development.
                    await _ensureHost(ct).ConfigureAwait(false);
                    using (var socket = new ClientWebSocket())
                    {
                        await socket.ConnectAsync(uri, ct).ConfigureAwait(false);
                        _socket = socket;
                        _connected = true;
                        failures = 0;
                        Log.Write("relay: connected to " + uri);
                        _postToPage("{\"__bridge\":\"open\"}");
                        await ReceiveLoopAsync(socket, ct).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    failures++;
                    Log.Write("relay: " + ex.Message);
                    if (failures == 3) _postToPage(StatusFrame(UiLanguage.T("ホストに接続できません: ", "Cannot connect to the host: ") + ex.Message + UiLanguage.T("  (ログ: ", "  (log: ") + AgentSettings.LogPath + ")"));
                }
                finally
                {
                    _connected = false;
                    _socket = null;
                    _postToPage("{\"__bridge\":\"closed\"}");
                }
                var delay = Math.Min(15000, 2000 * Math.Max(1, failures));
                try { await Task.Delay(delay, ct).ConfigureAwait(false); } catch (OperationCanceledException) { break; }
            }
        }

        /// <summary>
        /// Cheap prefix test so that large frames (chat messages carrying base64 attachments) are
        /// forwarded without ever being parsed inside the Office process.
        /// </summary>
        public static PageFrameKind ClassifyPageFrame(string text)
        {
            if (string.IsNullOrEmpty(text)) return PageFrameKind.Other;
            if (text.StartsWith("{\"__bridge\"", StringComparison.Ordinal)) return PageFrameKind.Bridge;
            if (text.StartsWith("{\"type\":\"hello\"", StringComparison.Ordinal)) return PageFrameKind.Hello;
            return PageFrameKind.Other;
        }

        /// <summary>The path itself when it exists, else its parent folder when that exists, else null.</summary>
        public static string ResolveOpenTarget(string path, Func<string, bool> fileExists, Func<string, bool> dirExists)
        {
            if (string.IsNullOrEmpty(path)) return null;
            if (fileExists(path) || dirExists(path)) return path;
            string parent;
            try { parent = Path.GetDirectoryName(path); }
            catch { return null; }
            return parent != null && dirExists(parent) ? parent : null;
        }

        /// <summary>Called for every message the page posts via window.chrome.webview.postMessage.</summary>
        public async Task OnPageMessageAsync(string text)
        {
            switch (ClassifyPageFrame(text))
            {
                case PageFrameKind.Other:
                    if (string.IsNullOrEmpty(text)) return;
                    await SendAsync(text).ConfigureAwait(false);
                    return;
                case PageFrameKind.Bridge:
                    HandleBridgeFrame(text);
                    return;
                case PageFrameKind.Hello:
                    JObject msg;
                    try { msg = JObject.Parse(text); }
                    catch { Log.Write("relay: malformed hello from page ignored"); return; }
                    msg["hostApp"] = _ops.HostApp;
                    msg["documentName"] = _ops.DocumentName;
                    msg["officeVersion"] = _ops.OfficeVersion;
                    await SendAsync(msg.ToString(Newtonsoft.Json.Formatting.None)).ConfigureAwait(false);
                    return;
            }
        }

        private void HandleBridgeFrame(string text)
        {
            JObject msg;
            try { msg = JObject.Parse(text); }
            catch { Log.Write("relay: malformed bridge frame ignored"); return; }
            var bridge = msg["__bridge"]?.ToString();
            switch (bridge)
            {
                case "ready":
                    // Page (re)loaded: tell it the current relay state.
                    _postToPage(_connected ? "{\"__bridge\":\"open\"}" : "{\"__bridge\":\"closed\"}");
                    break;
                case "openPath":
                    OpenPath(msg["path"]?.ToString());
                    break;
            }
        }

        /// <summary>Open a file or folder with the Windows default handler (used for the Copilot CLI settings).</summary>
        private static void OpenPath(string path)
        {
            var target = ResolveOpenTarget(path, File.Exists, Directory.Exists);
            if (target == null) { Log.Write("relay: openPath: nothing to open for " + path); return; }
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(target) { UseShellExecute = true }); }
            catch (Exception ex) { Log.Write("relay: openPath failed: " + ex.Message); }
        }

        private async Task ReceiveLoopAsync(ClientWebSocket socket, CancellationToken ct)
        {
            var buffer = new byte[64 * 1024];
            while (socket.State == WebSocketState.Open && !ct.IsCancellationRequested)
            {
                using (var ms = new MemoryStream())
                {
                    WebSocketReceiveResult result;
                    do
                    {
                        result = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), ct).ConfigureAwait(false);
                        if (result.MessageType == WebSocketMessageType.Close) return;
                        ms.Write(buffer, 0, result.Count);
                    } while (!result.EndOfMessage);

                    var text = Encoding.UTF8.GetString(ms.ToArray());
                    HandleHostMessage(text);
                }
            }
        }

        private void HandleHostMessage(string text)
        {
            JObject msg;
            try { msg = JObject.Parse(text); }
            catch { return; }

            if (msg["type"]?.ToString() != "office.request")
            {
                _postToPage(text);
                return;
            }

            var requestId = msg["requestId"]?.ToString();
            var op = msg["op"] as JObject ?? new JObject();
            // Run the COM work on the Office thread without blocking the receive loop.
            Task.Run(async () =>
            {
                JObject reply;
                var kind = op["kind"]?.ToString() ?? "?";
                Log.Write($"op {kind}: start");
                try
                {
                    var result = InvokeOnUiThread(() => _ops.Execute(op));
                    Log.Write($"op {kind}: ok");
                    reply = new JObject { ["type"] = "office.result", ["requestId"] = requestId, ["ok"] = true, ["result"] = result ?? JValue.CreateNull() };
                }
                catch (Exception ex)
                {
                    var inner = ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex;
                    Log.Write($"relay: op '{op["kind"]}' failed: {inner}");
                    reply = new JObject { ["type"] = "office.result", ["requestId"] = requestId, ["ok"] = false, ["error"] = Describe(inner) };
                }
                await SendAsync(reply.ToString(Newtonsoft.Json.Formatting.None)).ConfigureAwait(false);
            });
        }

        private JToken InvokeOnUiThread(Func<JToken> work)
        {
            if (_uiThread.IsDisposed) throw new OfficeOpException("Task pane is closed.");
            if (!_uiThread.InvokeRequired) return work();
            return (JToken)_uiThread.Invoke(work);
        }

        /// <summary>A protocol `status` frame the page shows in its status line.</summary>
        private static string StatusFrame(string message)
        {
            return new JObject { ["type"] = "status", ["message"] = message }.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string Describe(Exception ex)
        {
            if (ex is OfficeOpException) return ex.Message;
            if (ex is System.Runtime.InteropServices.COMException com)
                return $"Office COM error 0x{com.ErrorCode:X8}: {com.Message}";
            return ex.GetType().Name + ": " + ex.Message;
        }

        private async Task SendAsync(string text)
        {
            var socket = _socket;
            if (socket == null || socket.State != WebSocketState.Open)
            {
                Log.Write("relay: dropped frame (not connected)");
                return;
            }
            var bytes = Encoding.UTF8.GetBytes(text);
            await _sendGate.WaitAsync().ConfigureAwait(false);
            try
            {
                await socket.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Log.Write("relay: send failed: " + ex.Message);
            }
            finally
            {
                _sendGate.Release();
            }
        }

        public void Dispose()
        {
            try { _socket?.Abort(); } catch { /* ignore */ }
            _sendGate.Dispose();
        }
    }
}
