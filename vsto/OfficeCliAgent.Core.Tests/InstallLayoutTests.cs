using System;
using System.Collections.Generic;
using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class InstallLayoutTests
    {
        private static Func<string, bool> Files(params string[] paths)
        {
            var set = new HashSet<string>(paths, StringComparer.OrdinalIgnoreCase);
            return p => set.Contains(p);
        }

        [Fact]
        public void FindsRepositoryLayoutWalkingUp()
        {
            var exists = Files(@"C:\repo\packages\host\dist\cli.js", @"C:\repo\packages\ui\dist\taskpane.html");
            var r = InstallLayout.Find(@"C:\repo\vsto\OfficeCliAgent.Word\OfficeCliAgent.Word\bin\Debug", exists);
            Assert.NotNull(r);
            Assert.Equal(@"C:\repo\packages\host", r.HostDir);
            Assert.Equal(@"C:\repo\packages\ui\dist", r.UiDir);
            Assert.True(r.IsRepository);
        }

        [Fact]
        public void FindsInstalledLayoutWalkingUp()
        {
            var exists = Files(@"C:\inst\host\dist\cli.js", @"C:\inst\ui\taskpane.html");
            var r = InstallLayout.Find(@"C:\inst\addin\Word", exists);
            Assert.NotNull(r);
            Assert.Equal(@"C:\inst\host", r.HostDir);
            Assert.Equal(@"C:\inst\ui", r.UiDir);
            Assert.False(r.IsRepository);
        }

        [Fact]
        public void PrefersRepositoryOverInstalledAtSameLevel()
        {
            var exists = Files(
                @"C:\x\packages\host\dist\cli.js", @"C:\x\packages\ui\dist\taskpane.html",
                @"C:\x\host\dist\cli.js", @"C:\x\ui\taskpane.html");
            var r = InstallLayout.Find(@"C:\x\addin\Word", exists);
            Assert.NotNull(r);
            Assert.True(r.IsRepository);
            Assert.Equal(@"C:\x\packages\host", r.HostDir);
        }

        [Fact]
        public void ReturnsNullWhenNothingFound()
        {
            Assert.Null(InstallLayout.Find(@"C:\nothing\a\b", Files()));
        }

        [Fact]
        public void InstalledLayoutRequiresBothHostAndUi()
        {
            var exists = Files(@"C:\inst\host\dist\cli.js");
            Assert.Null(InstallLayout.Find(@"C:\inst\addin\Word", exists));
        }

        [Fact]
        public void StopsAfterMaxLevels()
        {
            var exists = Files(@"C:\inst\host\dist\cli.js", @"C:\inst\ui\taskpane.html");
            var start = @"C:\inst\l1\l2\l3\l4\l5\l6\l7\l8\l9\l10";
            Assert.Null(InstallLayout.Find(start, exists, 8));
            Assert.NotNull(InstallLayout.Find(start, exists, 11));
        }
    }
}
