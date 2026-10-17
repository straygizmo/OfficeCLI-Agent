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
            // The host outlives this Office process, so we must not hold pipes to it: a redirected
            // stdout/stderr keeps a thread blocked in a native read, and Office then hangs for tens of
            // seconds on exit while VSTO waits to unload the add-in. cmd.exe writes the output to a file instead.
            TrimHostLog();
            var psi = new ProcessStartInfo(Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe",
                $"/d /s /c \"\"{_settings.NodePath}\" {args} >> \"{AgentSettings.HostLogPath}\" 2>&1\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = _settings.HostDir,
            };

            Log.Write($"host: starting {_settings.NodePath} {args} (output: {AgentSettings.HostLogPath})");
            _process = Process.Start(psi);
            if (_process == null) throw new InvalidOperationException("Failed to start node.");

            var deadline = DateTime.UtcNow.AddSeconds(90);
            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();
                if (_process.HasExited)
                {
                    throw new InvalidOperationException($"Host exited with code {_process.ExitCode}.\n{ReadHostLogTail()}");
                }
                if (await IsListeningAsync(ct).ConfigureAwait(false)) return;
                await Task.Delay(500, ct).ConfigureAwait(false);
            }
            throw new TimeoutException("Host did not start listening within 90 seconds. See " + AgentSettings.HostLogPath);
        }

        private static void TrimHostLog()
        {
            try
            {
                var file = new FileInfo(AgentSettings.HostLogPath);
                if (file.Exists && file.Length > 1024 * 1024) file.Delete();
            }
            catch { /* a running host may still hold it; keep appending */ }
        }

        private static string ReadHostLogTail()
        {
            try
            {
                using (var stream = new FileStream(AgentSettings.HostLogPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    stream.Seek(Math.Max(0, stream.Length - 4000), SeekOrigin.Begin);
                    using (var reader = new StreamReader(stream, Encoding.UTF8)) return reader.ReadToEnd();
                }
            }
            catch { return ""; }
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
