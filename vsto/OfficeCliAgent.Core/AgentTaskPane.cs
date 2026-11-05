using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// WinForms control hosted in the VSTO custom task pane. Shows the React chat UI
    /// (packages/ui/dist) in WebView2 and exposes the page's postMessage channel.
    /// </summary>
    public sealed class AgentTaskPane : UserControl
    {
        private const string VirtualHost = "officecli-agent.local";
        private const int InitAttempts = 3;
        private static Task<CoreWebView2Environment> _sharedEnvironment;
        private WebView2 _web;
        private readonly Label _status = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Padding = new Padding(12), AutoSize = false };
        private bool _ready;

        /// <summary>Raised on the UI thread with the raw string the page posted.</summary>
        public event Action<string> MessageFromPage;

        public AgentTaskPane()
        {
            Dock = DockStyle.Fill;
            Controls.Add(_status);
            _web = NewWebView();
            _status.Text = UiLanguage.T("起動中…", "Starting…");
        }

        private WebView2 NewWebView()
        {
            var web = new WebView2 { Dock = DockStyle.Fill, Visible = false };
            Controls.Add(web);
            return web;
        }

        /// <summary>
        /// One WebView2 environment per Office process, shared by every pane (one per document window).
        /// Every WebView2 in a process must be created with identical browser options, and Office
        /// changes WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS in its own process while it runs
        /// (https://github.com/MicrosoftEdge/WebView2Feedback/issues/5770): an environment created
        /// later picks up different arguments and its controllers fail with 0x8007139F.
        /// </summary>
        private static Task<CoreWebView2Environment> SharedEnvironmentAsync()
        {
            var task = _sharedEnvironment;
            if (task == null || task.IsFaulted || task.IsCanceled)
            {
                Log.Write("webview: creating the environment (extra browser arguments: "
                    + (Environment.GetEnvironmentVariable("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS") ?? "none") + ")");
                task = _sharedEnvironment = CoreWebView2Environment.CreateAsync(null, AgentSettings.WebViewDataDir);
            }
            return task;
        }

        public async Task InitializeAsync(AgentSettings settings, string hostApp)
        {
            var entry = settings.UiEntryPath;
            if (entry == null || !File.Exists(entry))
            {
                ShowError(UiLanguage.T(
                    "チャット UI が見つかりません。OfficeCLI Agent を再インストールするか、開発時はリポジトリで pnpm build を実行して " + AgentSettings.ConfigPath + " の UiDir を設定してください。",
                    "Chat UI not found. Reinstall OfficeCLI Agent, or (for development) run pnpm build and set UiDir in " + AgentSettings.ConfigPath + ".")
                    + "\n\n" + (entry ?? "(UiDir unset)"));
                return;
            }

            try
            {
                var env = await SharedEnvironmentAsync();
                // Office may re-create the pane's window while attaching it to a document window;
                // a controller created against a handle that went away fails, so try again with a
                // fresh control a couple of times before showing the error.
                for (var attempt = 1; ; attempt++)
                {
                    try
                    {
                        await _web.EnsureCoreWebView2Async(env);
                        break;
                    }
                    catch (Exception ex) when (attempt < InitAttempts && !IsDisposed)
                    {
                        Log.Write($"webview: init attempt {attempt} failed ({ex.Message.Trim()}); retrying with a new control");
                        _web.Dispose();
                        _web = NewWebView();
                        await Task.Delay(300 * attempt);
                    }
                }
            }
            catch (Exception ex)
            {
                if (IsDisposed) return;
                ShowError(UiLanguage.T(
                    "WebView2 を初期化できません。Microsoft Edge WebView2 ランタイムをインストールしてください。",
                    "Could not initialize WebView2. Install the Microsoft Edge WebView2 Runtime.")
                    + "\n\n" + ex.Message);
                Log.Write("webview: " + ex);
                return;
            }

            var core = _web.CoreWebView2;
            // Files dragged from Explorer must reach the page's drop handler (attachments).
            _web.AllowExternalDrop = true;
            core.SetVirtualHostNameToFolderMapping(VirtualHost, settings.UiDir, CoreWebView2HostResourceAccessKind.Allow);
            core.Settings.AreDefaultContextMenusEnabled = false;
            core.Settings.IsStatusBarEnabled = false;
            core.WebMessageReceived += OnWebMessageReceived;
            core.NewWindowRequested += (_, e) =>
            {
                // Open external links in the default browser instead of a WebView2 popup.
                e.Handled = true;
                try { System.Diagnostics.Process.Start(e.Uri); } catch { /* ignore */ }
            };

            _web.Source = new Uri($"https://{VirtualHost}/taskpane.html?hostApp={Uri.EscapeDataString(hostApp ?? "")}&lang={UiLanguage.Code}");
            _web.Visible = true;
            _status.Visible = false;
            _ready = true;
        }

        public void ShowError(string message)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => ShowError(message))); return; }
            _status.Text = message;
            _status.Visible = true;
            _web.Visible = false;
        }

        /// <summary>Deliver a JSON frame to the page (thread-safe).</summary>
        public void PostToPage(string json)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { try { BeginInvoke((Action)(() => PostToPage(json))); } catch { /* disposed */ } return; }
            if (!_ready || _web.CoreWebView2 == null) return;
            try { _web.CoreWebView2.PostWebMessageAsString(json); }
            catch (Exception ex) { Log.Write("webview: post failed: " + ex.Message); }
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string text;
            try { text = e.TryGetWebMessageAsString(); }          // page posted a string (chat frames)
            catch { text = e.WebMessageAsJson; }                    // page posted an object ({"__bridge":"ready"})
            MessageFromPage?.Invoke(text);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _web.Dispose();
            base.Dispose(disposing);
        }
    }
}
