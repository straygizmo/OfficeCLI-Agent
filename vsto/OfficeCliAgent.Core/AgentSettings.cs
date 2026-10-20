using System;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Add-in side settings. Stored in %LOCALAPPDATA%\OfficeCliAgent\addin.json;
    /// environment variables OFFICECLI_AGENT_PORT / _HOST_DIR / _UI_DIR / _NODE override it.
    /// Unset host/UI directories are resolved by walking up from the add-in assembly
    /// until either a repo checkout (packages/host/dist/cli.js, packages/ui/dist/taskpane.html)
    /// or an installed layout (host/dist/cli.js, ui/taskpane.html next to addin/) is found.
    /// See <see cref="InstallLayout"/>.
    /// </summary>
    public sealed class AgentSettings
    {
        public int Port { get; set; } = 4320;

        /// <summary>Node executable ("node" = resolve from PATH).</summary>
        public string NodePath { get; set; } = "node";

        /// <summary>Folder containing the built host (dist/cli.js + node_modules).</summary>
        public string HostDir { get; set; }

        /// <summary>Folder containing the built chat UI (taskpane.html + assets/).</summary>
        public string UiDir { get; set; }

        /// <summary>Start the host process automatically when nothing listens on Port.</summary>
        public bool AutoStartHost { get; set; } = true;

        /// <summary>Seconds the host stays alive with no connected add-in before exiting (0 = never).</summary>
        public int HostIdleExitSeconds { get; set; } = 900;

        public static string DataDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeCliAgent");
        public static string ConfigPath => Path.Combine(DataDir, "addin.json");
        public static string LogPath => Path.Combine(DataDir, "addin.log");
        public static string HostLogPath => Path.Combine(DataDir, "host.log");
        public static string WebViewDataDir => Path.Combine(DataDir, "WebView2");

        public static AgentSettings Load()
        {
            AgentSettings s = null;
            try
            {
                if (File.Exists(ConfigPath))
                    s = JsonConvert.DeserializeObject<AgentSettings>(File.ReadAllText(ConfigPath));
            }
            catch (Exception ex)
            {
                Log.Write("settings: failed to read " + ConfigPath + ": " + ex.Message);
            }
            s = s ?? new AgentSettings();

            var env = Environment.GetEnvironmentVariable("OFFICECLI_AGENT_PORT");
            if (int.TryParse(env, out var port)) s.Port = port;
            s.HostDir = Environment.GetEnvironmentVariable("OFFICECLI_AGENT_HOST_DIR") ?? s.HostDir;
            s.UiDir = Environment.GetEnvironmentVariable("OFFICECLI_AGENT_UI_DIR") ?? s.UiDir;
            s.NodePath = Environment.GetEnvironmentVariable("OFFICECLI_AGENT_NODE") ?? s.NodePath;

            // The first-run file never contains resolved directories: resolution happens on
            // every load, so the same file keeps working after a reinstall or an uninstall.
            Directory.CreateDirectory(DataDir);
            if (!File.Exists(ConfigPath))
            {
                try { File.WriteAllText(ConfigPath, SerializeDefaults(s)); }
                catch (Exception ex) { Log.Write("settings: failed to write defaults: " + ex.Message); }
            }

            var layout = FindLayout();
            if (string.IsNullOrEmpty(s.HostDir) && layout != null) s.HostDir = layout.HostDir;
            if (string.IsNullOrEmpty(s.UiDir) && layout != null) s.UiDir = layout.UiDir;
            Log.Write("settings: layout " + (layout == null ? "none" : layout.IsRepository ? "repository" : "installed")
                + " host=" + (s.HostDir ?? "(unset)") + " ui=" + (s.UiDir ?? "(unset)"));
            return s;
        }

        /// <summary>JSON written on first run: user-settable fields only, directories left unset.</summary>
        public static string SerializeDefaults(AgentSettings s)
        {
            var copy = new AgentSettings
            {
                Port = s.Port,
                NodePath = s.NodePath,
                HostDir = null,
                UiDir = null,
                AutoStartHost = s.AutoStartHost,
                HostIdleExitSeconds = s.HostIdleExitSeconds,
            };
            return JsonConvert.SerializeObject(copy, Formatting.Indented);
        }

        [JsonIgnore]
        public string HostCliPath => string.IsNullOrEmpty(HostDir) ? null : Path.Combine(HostDir, "dist", "cli.js");
        [JsonIgnore]
        public string UiEntryPath => string.IsNullOrEmpty(UiDir) ? null : Path.Combine(UiDir, "taskpane.html");

        /// <summary>Walk up from the add-in assembly looking for a repository or installed layout.</summary>
        private static InstallLayout.Result FindLayout()
        {
            try
            {
                var dir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                return InstallLayout.Find(dir, File.Exists);
            }
            catch { return null; }
        }
    }

    public static class Log
    {
        private static readonly object Gate = new object();

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(AgentSettings.DataDir);
                    File.AppendAllText(AgentSettings.LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine);
                }
            }
            catch { /* never throw from logging */ }
        }
    }
}
