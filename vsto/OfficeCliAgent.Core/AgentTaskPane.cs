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
        private readonly WebView2 _web = new WebView2 { Dock = DockStyle.Fill };
        private readonly Label _status = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Padding = new Padding(12), AutoSize = false };
        private bool _ready;

        /// <summary>Raised on the UI thread with the raw string the page posted.</summary>
        public event Action<string> MessageFromPage;

        public AgentTaskPane()
        {
            Dock = DockStyle.Fill;
            Controls.Add(_status);
            Controls.Add(_web);
            _web.Visible = false;
            _status.Text = UiLanguage.T("起動中…", "Starting…");
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
                var env = await CoreWebView2Environment.CreateAsync(null, AgentSettings.WebViewDataDir);
                await _web.EnsureCoreWebView2Async(env);
            }
            catch (Exception ex)
            {
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
