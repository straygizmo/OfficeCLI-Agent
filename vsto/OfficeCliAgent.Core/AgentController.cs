using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Ties the pieces together for one Office application instance:
    /// task-pane control, host process, and the relay. Create it in ThisAddIn_Startup
    /// and hand <see cref="Pane"/> to a VSTO CustomTaskPane.
    /// </summary>
    public sealed class AgentController : IDisposable
    {
        private readonly IOfficeOps _ops;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        private AgentSettings _settings;
        private HostProcessManager _process;
        private HostRelay _relay;
        private bool _started;

        public AgentTaskPane Pane { get; } = new AgentTaskPane();

        public AgentController(IOfficeOps ops)
        {
            _ops = ops;
            // WebView2 needs a window handle; start once the pane is actually shown.
            Pane.HandleCreated += (_, __) => Start();
        }

        public void Start()
        {
            if (_started) return;
            _started = true;
            _ = StartAsync();
        }

        private async Task StartAsync()
        {
            try
            {
                _settings = AgentSettings.Load();
                Log.Write($"controller: start ({_ops.HostApp}, port {_settings.Port}, host {_settings.HostDir}, ui {_settings.UiDir})");

                await Pane.InitializeAsync(_settings, _ops.HostApp);

                _process = new HostProcessManager(_settings);
                _relay = new HostRelay(_settings, _ops, Pane, Pane.PostToPage, EnsureHostAsync);
                Pane.MessageFromPage += text => _ = _relay.OnPageMessageAsync(text);

                // The relay starts the host itself before each connection attempt (see EnsureHostAsync),
                // so a host that exits on idle or is killed during development comes back automatically.
                _ = Task.Run(() => _relay.RunAsync(_cts.Token));
            }
            catch (Exception ex)
            {
                Log.Write("controller: " + ex);
                Pane.ShowError(UiLanguage.T("初期化に失敗しました。", "Initialization failed.") + "\n\n" + ex.Message);
            }
        }

        private readonly SemaphoreSlim _ensureGate = new SemaphoreSlim(1, 1);

        private async Task EnsureHostAsync(CancellationToken ct)
        {
            await _ensureGate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                await _process.EnsureRunningAsync(ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Log.Write("controller: host start failed: " + ex.Message);
                throw;
            }
            finally
            {
                _ensureGate.Release();
            }
        }

        public void Dispose()
        {
            _cts.Cancel();
            _relay?.Dispose();
            _process?.Dispose();
            Pane.Dispose();
            _ensureGate.Dispose();
            _cts.Dispose();
        }
    }
}
