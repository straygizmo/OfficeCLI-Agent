using System;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Makes sure the Node host (packages/host) is listening on the configured port,
    /// starting it as a hidden child process when needed. The host is shared by all
    /// Office applications and exits itself after being idle (see --idle-exit).
    /// </summary>
    public sealed class HostProcessManager : IDisposable
    {
        private readonly AgentSettings _settings;
        private Process _process;

        public HostProcessManager(AgentSettings settings)
        {
            _settings = settings;
        }

        public async Task EnsureRunningAsync(CancellationToken ct)
        {
            if (await IsListeningAsync(ct).ConfigureAwait(false)) return;
            if (!_settings.AutoStartHost)
                throw new InvalidOperationException($"No host is listening on port {_settings.Port} and AutoStartHost is disabled.");

            var cli = _settings.HostCliPath;
            if (cli == null || !File.Exists(cli))
                throw new FileNotFoundException(
                    "Host not found. Build the repository (pnpm build) or set HostDir in " + AgentSettings.ConfigPath, cli ?? "(unset)");

            var args = $"\"{cli}\" --port {_settings.Port} --no-tls --idle-exit {_settings.HostIdleExitSeconds}";
            var psi = new ProcessStartInfo(_settings.NodePath, args)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = _settings.HostDir,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            var stderr = new StringBuilder();
            Log.Write($"host: starting {_settings.NodePath} {args}");
            _process = Process.Start(psi);
            if (_process == null) throw new InvalidOperationException("Failed to start node.");
            _process.OutputDataReceived += (_, e) => { if (e.Data != null) Log.Write("host> " + e.Data); };
            _process.ErrorDataReceived += (_, e) =>
            {
                if (e.Data == null) return;
                Log.Write("host! " + e.Data);
                lock (stderr) { if (stderr.Length < 4000) stderr.AppendLine(e.Data); }
            };
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            var deadline = DateTime.UtcNow.AddSeconds(90);
            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();
                if (_process.HasExited)
                {
                    string tail; lock (stderr) tail = stderr.ToString();
                    throw new InvalidOperationException($"Host exited with code {_process.ExitCode}.\n{tail}");
                }
                if (await IsListeningAsync(ct).ConfigureAwait(false)) return;
                await Task.Delay(500, ct).ConfigureAwait(false);
            }
            throw new TimeoutException("Host did not start listening within 90 seconds. See " + AgentSettings.LogPath);
        }

        public async Task<bool> IsListeningAsync(CancellationToken ct)
        {
            using (var client = new TcpClient())
            {
                try
                {
                    var connect = client.ConnectAsync("127.0.0.1", _settings.Port);
                    var done = await Task.WhenAny(connect, Task.Delay(1500, ct)).ConfigureAwait(false);
                    if (done != connect) return false;
                    await connect.ConfigureAwait(false);
                    return client.Connected;
                }
                catch
                {
                    return false;
                }
            }
        }

        /// <summary>We never kill the host here: other Office apps may share it. It exits on idle.</summary>
        public void Dispose()
        {
            _process?.Dispose();
            _process = null;
        }
    }
}
