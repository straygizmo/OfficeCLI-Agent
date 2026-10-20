using System;
using System.IO;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Locates the Node host and the chat UI relative to the add-in assembly.
    /// Two layouts are recognised, checked at every directory walking up from the start:
    ///   repository: &lt;root&gt;\packages\host\dist\cli.js and &lt;root&gt;\packages\ui\dist\taskpane.html
    ///   installed:  &lt;root&gt;\host\dist\cli.js and &lt;root&gt;\ui\taskpane.html (next to &lt;root&gt;\addin\&lt;App&gt;)
    /// The repository layout wins when both exist at the same level.
    /// </summary>
    public static class InstallLayout
    {
        public sealed class Result
        {
            public string HostDir;
            public string UiDir;
            public bool IsRepository;
        }

        public static Result Find(string startDir, Func<string, bool> fileExists, int maxLevels = 8)
        {
            var dir = startDir;
            for (var i = 0; i < maxLevels && !string.IsNullOrEmpty(dir); i++)
            {
                if (fileExists(Path.Combine(dir, "packages", "host", "dist", "cli.js")))
                {
                    return new Result
                    {
                        HostDir = Path.Combine(dir, "packages", "host"),
                        UiDir = Path.Combine(dir, "packages", "ui", "dist"),
                        IsRepository = true,
                    };
                }
                if (fileExists(Path.Combine(dir, "host", "dist", "cli.js")) &&
                    fileExists(Path.Combine(dir, "ui", "taskpane.html")))
                {
                    return new Result
                    {
                        HostDir = Path.Combine(dir, "host"),
                        UiDir = Path.Combine(dir, "ui"),
                        IsRepository = false,
                    };
                }
                dir = Path.GetDirectoryName(dir);
            }
            return null;
        }
    }
}
