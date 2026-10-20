using Newtonsoft.Json.Linq;
using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class AgentSettingsTests
    {
        private static string Json() =>
            AgentSettings.SerializeDefaults(new AgentSettings { HostDir = @"C:\resolved\host", UiDir = @"C:\resolved\ui" });

        [Fact]
        public void DefaultsJsonLeavesDirectoriesUnset()
        {
            var o = JObject.Parse(Json());
            Assert.Equal(JTokenType.Null, o["HostDir"].Type);
            Assert.Equal(JTokenType.Null, o["UiDir"].Type);
            Assert.Equal(4320, (int)o["Port"]);
            Assert.Equal("node", (string)o["NodePath"]);
            Assert.True((bool)o["AutoStartHost"]);
            Assert.Equal(900, (int)o["HostIdleExitSeconds"]);
        }

        [Fact]
        public void DefaultsJsonHasNoDerivedPaths()
        {
            var json = Json();
            Assert.DoesNotContain("HostCliPath", json);
            Assert.DoesNotContain("UiEntryPath", json);
        }
    }
}
