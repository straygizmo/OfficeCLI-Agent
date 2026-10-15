using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;
using OfficeCliAgent.Core;
using Interop = Microsoft.Office.Interop.Word;

namespace OfficeCliAgent.Word
{
    /// <summary>Protocol OfficeOp implementation for Word (COM).</summary>
    public sealed class WordOps : IOfficeOps
    {
        private const int PreviewChars = 6000;
        private readonly Interop.Application _app;

        public WordOps(Interop.Application app) { _app = app; }

        public string HostApp => "Word";
        public string DocumentName => _app.Documents.Count > 0 ? _app.ActiveDocument.Name : null;
        public string OfficeVersion => "Word " + _app.Version + " build " + _app.Build;

        private Interop.Document Doc
        {
            get
            {
                if (_app.Documents.Count == 0) throw new OfficeOpException("No document is open in Word.");
                return _app.ActiveDocument;
            }
        }

        public JToken Execute(JObject op)
        {
            var kind = Ops.Str(op, "kind");
            switch (kind)
            {
                case "getContext": return GetContext();
                case "exportFile": return ExportFile(Ops.Str(op, "targetPath"));
                case "importFile": return ImportFile(Ops.Str(op, "path"), Ops.Options(op));
                case "word.getSelection": return GetSelection();
                case "word.insertText": return InsertText(Ops.Str(op, "text", ""), Ops.Str(op, "location", "Selection"), Ops.Str(op, "style"));
                case "word.replaceText": return ReplaceText(Ops.Str(op, "find"), Ops.Str(op, "replace", ""), Ops.Bool(op, "matchCase"), Ops.Bool(op, "replaceAll", true));
                default: throw new OfficeOpException($"Operation '{kind}' is not supported in Word.");
            }
        }

        // ------------------------------------------------------------------

        private JToken GetContext()
        {
            var doc = Doc;
            var sb = new StringBuilder();
            var saved = doc.FullName.Contains("\\") || doc.FullName.Contains("/");
            sb.AppendLine("Host: Word");
            sb.AppendLine("Document: " + doc.Name + (saved ? " (" + doc.FullName + ")" : " (unsaved)"));
            sb.AppendLine($"Paragraphs: {doc.Paragraphs.Count}, characters: {doc.Characters.Count}, pages: {SafePages(doc)}");

            var sel = _app.Selection;
            sb.AppendLine("Selection: " + JValue.CreateString(Ops.Clip(sel.Text, 500)).ToString());

            var headings = new List<string>();
            var i = 0;
            foreach (Interop.Paragraph p in doc.Paragraphs)
            {
                if (++i > 3000 || headings.Count >= 60) break;
                string styleName;
                try { styleName = ((Interop.Style)p.get_Style()).NameLocal; } catch { continue; }
                if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) || styleName.StartsWith("見出し") || styleName.Equals("Title", StringComparison.OrdinalIgnoreCase) || styleName == "表題")
                    headings.Add($"  [{styleName}] {Ops.Clip(p.Range.Text.TrimEnd('\r', '\a'), 120)}");
            }
            sb.AppendLine(headings.Count > 0 ? "Headings:\n" + string.Join("\n", headings) : "Headings: (none)");
            sb.AppendLine("Text preview:");
            sb.Append(Ops.Clip(doc.Content.Text, PreviewChars));

            return Ops.Context(HostApp, doc.Name, saved ? doc.FullName : null, sb.ToString());
        }

        private static int SafePages(Interop.Document doc)
        {
            try { return doc.ComputeStatistics(Interop.WdStatistic.wdStatisticPages); } catch { return -1; }
        }

        private JToken ExportFile(string targetPath)
        {
            if (string.IsNullOrEmpty(targetPath)) throw new OfficeOpException("targetPath is required.");
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
            if (File.Exists(targetPath)) File.Delete(targetPath);
            // ExportFragment writes the range (whole body here) as a standalone .docx without
            // changing the document's own path. Headers/footers and section layout travel with it.
            Doc.Content.ExportFragment(targetPath, Interop.WdSaveFormat.wdFormatXMLDocument);
            return new JObject { ["path"] = targetPath };
        }

        private JToken ImportFile(string path, JObject options)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) throw new OfficeOpException("File not found: " + path);
            var location = Ops.Str(options, "location", "End");
            var doc = Doc;
            Interop.Range target;
            switch (location)
            {
                case "Selection":
                    target = _app.Selection.Range;
                    break;
                case "Start":
                    target = doc.Content;
                    target.Collapse(Interop.WdCollapseDirection.wdCollapseStart);
                    break;
                case "Replace":
                    doc.Content.Delete();
                    target = doc.Content;
                    break;
                default:
                    target = doc.Content;
                    target.Collapse(Interop.WdCollapseDirection.wdCollapseEnd);
                    break;
            }
            var before = doc.Characters.Count;
            target.InsertFile(path, ConfirmConversions: false, Link: false, Attachment: false);
            var after = doc.Characters.Count;
            return $"Inserted {Path.GetFileName(path)} ({location}); characters {before} -> {after}.";
        }

        private JToken GetSelection()
        {
            var sel = _app.Selection;
            return new JObject { ["text"] = sel.Text, ["start"] = sel.Start, ["end"] = sel.End, ["type"] = sel.Type.ToString() };
        }

        private JToken InsertText(string text, string location, string style)
        {
            var doc = Doc;
            Interop.Range range;
            switch (location)
            {
                case "Start":
                    range = doc.Content;
                    range.Collapse(Interop.WdCollapseDirection.wdCollapseStart);
                    break;
                case "End":
                    range = doc.Content;
                    range.Collapse(Interop.WdCollapseDirection.wdCollapseEnd);
                    break;
                default:
                    range = _app.Selection.Range;
                    break;
            }
            range.Text = text.Replace("\r\n", "\n").Replace("\n", "\r");
            if (!string.IsNullOrEmpty(style))
            {
                try
                {
                    object st = style;
                    range.set_Style(ref st);
                }
                catch (Exception ex)
                {
                    return $"Inserted {text.Length} characters, but style '{style}' could not be applied: {ex.Message}";
                }
            }
            return $"Inserted {text.Length} characters at {location}.";
        }

        private JToken ReplaceText(string find, string replace, bool matchCase, bool replaceAll)
        {
            if (string.IsNullOrEmpty(find)) throw new OfficeOpException("find is required.");
            var doc = Doc;

            // Count occurrences first so the agent gets a number back.
            var count = 0;
            var probe = doc.Content;
            var pf = probe.Find;
            pf.ClearFormatting();
            pf.Text = find;
            pf.MatchCase = matchCase;
            pf.Forward = true;
            pf.Wrap = Interop.WdFindWrap.wdFindStop;
            while (count < 10000 && pf.Execute())
            {
                count++;
                probe.Collapse(Interop.WdCollapseDirection.wdCollapseEnd);
            }
            if (count == 0) return new JObject { ["replaced"] = 0 };

            var range = doc.Content;
            var f = range.Find;
            f.ClearFormatting();
            f.Replacement.ClearFormatting();
            f.Text = find;
            f.Replacement.Text = replace;
            f.MatchCase = matchCase;
            f.Forward = true;
            f.Wrap = Interop.WdFindWrap.wdFindContinue;
            f.Execute(Replace: replaceAll ? Interop.WdReplace.wdReplaceAll : Interop.WdReplace.wdReplaceOne);
            return new JObject { ["replaced"] = replaceAll ? count : 1, ["matches"] = count };
        }
    }
}
