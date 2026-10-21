using Newtonsoft.Json.Linq;
using OfficeCliAgent.Core;
using Xunit;

namespace OfficeCliAgent.Core.Tests
{
    public class OpsOpenDocumentTests
    {
        [Theory]
        [InlineData("report.docx")]
        [InlineData("REPORT.DOCX")]
        [InlineData("report")]
        [InlineData(@"C:\docs\report.docx")]
        public void NameMatchesAcceptsNamePathAndStem(string wanted)
        {
            Assert.True(Ops.NameMatches(wanted, "report.docx", @"C:\docs\report.docx"));
        }

        [Theory]
        [InlineData("other.docx")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(@"C:\elsewhere\report.docx")]
        public void NameMatchesRejectsOtherDocuments(string wanted)
        {
            Assert.False(Ops.NameMatches(wanted, "report.docx", @"C:\docs\report.docx"));
        }

        [Fact]
        public void NameMatchesHandlesUnsavedDocuments()
        {
            // Office reports just the name as FullName until the document is saved.
            Assert.True(Ops.NameMatches("Document1", "Document1", "Document1"));
            Assert.False(Ops.NameMatches("Document2", "Document1", "Document1"));
        }

        [Fact]
        public void OpenDocumentReportsPathOnlyWhenSaved()
        {
            var saved = Ops.OpenDocument("report.docx", @"C:\docs\report.docx", true);
            Assert.Equal(@"C:\docs\report.docx", (string)saved["path"]);
            Assert.True((bool)saved["saved"]);
            Assert.True((bool)saved["active"]);

            var unsaved = Ops.OpenDocument("Document1", "Document1", false);
            Assert.Equal(JTokenType.Null, unsaved["path"].Type);
            Assert.False((bool)unsaved["saved"]);
            Assert.False((bool)unsaved["active"]);
        }
    }
}
