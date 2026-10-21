using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class HostRelayTests
    {
        [Theory]
        [InlineData("{\"__bridge\":\"ready\"}", PageFrameKind.Bridge)]
        [InlineData("{\"__bridge\":\"openPath\",\"path\":\"C:\\\\x\"}", PageFrameKind.Bridge)]
        [InlineData("{\"type\":\"hello\",\"protocolVersion\":1}", PageFrameKind.Hello)]
        [InlineData("{\"type\":\"chat.send\",\"text\":\"hi\",\"attachments\":[{\"dataBase64\":\"AAAA\"}]}", PageFrameKind.Other)]
        [InlineData("not json at all", PageFrameKind.Other)]
        [InlineData("", PageFrameKind.Other)]
        public void ClassifyPageFrameLooksOnlyAtThePrefix(string text, PageFrameKind expected)
        {
            Assert.Equal(expected, HostRelay.ClassifyPageFrame(text));
        }

        [Fact]
        public void ResolveOpenTargetFallsBackToTheParentFolder()
        {
            Assert.Equal(@"C:\c\mcp-config.json", HostRelay.ResolveOpenTarget(@"C:\c\mcp-config.json", p => p == @"C:\c\mcp-config.json", _ => false));
            Assert.Equal(@"C:\c", HostRelay.ResolveOpenTarget(@"C:\c\mcp-config.json", _ => false, d => d == @"C:\c"));
            Assert.Null(HostRelay.ResolveOpenTarget(@"C:\nope\x.json", _ => false, _ => false));
        }
    }
}
