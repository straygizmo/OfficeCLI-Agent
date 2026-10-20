using System;
using System.IO;
using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class HostProcessManagerTests
    {
        [Fact]
        public void FindOnPathReturnsFirstDirectoryContainingTheFile()
        {
            var a = Path.Combine(Path.GetTempPath(), "ocagent-path-a-" + Guid.NewGuid().ToString("N"));
            var b = Path.Combine(Path.GetTempPath(), "ocagent-path-b-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(a);
            Directory.CreateDirectory(b);
            try
            {
                File.WriteAllText(Path.Combine(b, "node.exe"), "");
                var found = HostProcessManager.FindOnPath("node.exe", a + ";" + b + @";C:\does\not\exist");
                Assert.Equal(Path.Combine(b, "node.exe"), found);
            }
            finally
            {
                Directory.Delete(a, true);
                Directory.Delete(b, true);
            }
        }

        [Fact]
        public void FindOnPathReturnsNullWhenAbsent()
        {
            var name = "surely-missing-" + Guid.NewGuid().ToString("N") + ".exe";
            Assert.Null(HostProcessManager.FindOnPath(name, @"C:\;C:\does\not\exist"));
        }
    }
}
