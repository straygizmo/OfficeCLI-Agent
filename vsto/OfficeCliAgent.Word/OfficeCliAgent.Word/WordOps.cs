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
                case "exportFile": return ExportFile(Ops.Str(op, "targetPath"), Ops.Str(op, "documentName"));
                case "listOpenDocuments": return ListOpenDocuments();
                case "importFile": return ImportFile(Ops.Str(op, "path"), Ops.Options(op));
                case "insertImage": return InsertImage(Ops.Str(op, "path"), Ops.Options(op));
                case "word.getSelection": return GetSelection();
                case "word.insertText": return InsertText(Ops.Str(op, "text", ""), Ops.Str(op, "location", "Selection"), Ops.Str(op, "style"), Ops.Str(op, "list"));
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

        private JToken ListOpenDocuments()
        {
            var activeName = _app.Documents.Count > 0 ? _app.ActiveDocument.FullName : null;
            var list = new JArray();
            foreach (Interop.Document d in _app.Documents)
                list.Add(Ops.OpenDocument(d.Name, d.FullName, d.FullName == activeName));
            return list;
        }

        /// <summary>The active document, or the open document the agent named (listOpenDocuments names or paths).</summary>
        private Interop.Document FindDocument(string documentName)
        {
            if (string.IsNullOrEmpty(documentName)) return Doc;
            foreach (Interop.Document d in _app.Documents)
                if (Ops.NameMatches(documentName, d.Name, d.FullName)) return d;
            throw new OfficeOpException("No open Word document named '" + documentName + "'. Use office_list_open_documents to see the open files.");
        }

        private JToken ExportFile(string targetPath, string documentName)
        {
            if (string.IsNullOrEmpty(targetPath)) throw new OfficeOpException("targetPath is required.");
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
            if (File.Exists(targetPath)) File.Delete(targetPath);
            // ExportFragment writes the range (whole body here) as a standalone .docx without
            // changing the document's own path. Headers/footers and section layout travel with it.
            FindDocument(documentName).Content.ExportFragment(targetPath, Interop.WdSaveFormat.wdFormatXMLDocument);
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

        private JToken InsertImage(string path, JObject options)
        {
            Ops.RequireImageFile(path, options);
            var location = Ops.Str(options, "location", "Selection");
            var range = TargetRange(location);

            // Inline picture, embedded in the document. A non-collapsed range is replaced.
            var pic = Doc.InlineShapes.AddPicture(path, LinkToFile: false, SaveWithDocument: true, Range: range);
            Ops.FitSize(pic.Width, pic.Height, Ops.Num(options, "width"), Ops.Num(options, "height"), out var w, out var h);
            if (w != pic.Width || h != pic.Height)
            {
                pic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                pic.Width = w;
                pic.Height = h;
                pic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            var alt = Ops.Str(options, "altText");
            if (!string.IsNullOrEmpty(alt)) pic.AlternativeText = alt;
            return $"Inserted picture {Path.GetFileName(path)} at {location} ({pic.Width:0.#} x {pic.Height:0.#} pt).";
        }

        /// <summary>Selection range, or a collapsed range at the start/end of the body.</summary>
        private Interop.Range TargetRange(string location)
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
            return range;
        }

        private JToken GetSelection()
        {
            var sel = _app.Selection;
            return new JObject { ["text"] = sel.Text, ["start"] = sel.Start, ["end"] = sel.End, ["type"] = sel.Type.ToString() };
        }

        private JToken InsertText(string text, string location, string style, string list)
        {
            var doc = Doc;
            var body = (text ?? "").Replace("\r\n", "\n").Replace("\n", "\r");
            Interop.Range inserted;
            object inherited = null; // style for the new paragraphs when the caller names none
            var newParagraphs = false;

            if (location == "End")
            {
                // Always start new paragraph(s): writing into the last paragraph would join the text to it
                // and a style meant for the new text would restyle that whole paragraph.
                body = body.TrimEnd('\r');
                var last = doc.Paragraphs.Last;
                var lastIsEmpty = (last.Range.Text ?? "").Trim('\r', '\n', '\a', ' ', '\t', '　').Length == 0;
                var pos = doc.Content.End - 1; // just before the final paragraph mark
                var at = doc.Range(pos, pos);
                if (lastIsEmpty)
                {
                    at.Text = body;
                    inserted = doc.Range(at.Start, at.End);
                }
                else
                {
                    // What pressing Enter at the end would give: the style that follows the last
                    // paragraph's style (body text after a heading), not a copy of the heading.
                    try { inherited = ((Interop.Style)last.get_Style()).get_NextParagraphStyle(); } catch { }
                    at.Text = "\r" + body;
                    inserted = doc.Range(at.Start + 1, at.End);
                    newParagraphs = true;
                }
            }
            else if (location == "Start")
            {
                body = body.TrimEnd('\r') + "\r";
                var at = doc.Content;
                at.Collapse(Interop.WdCollapseDirection.wdCollapseStart);
                at.Text = body;
                inserted = doc.Range(at.Start, at.End - 1);
                newParagraphs = true;
            }
            else
            {
                inserted = _app.Selection.Range;
                inserted.Text = body;
            }

            var notes = new List<string>();
            if (!string.IsNullOrEmpty(style))
            {
                try
                {
                    object st = style;
                    inserted.set_Style(ref st);
                }
                catch (Exception ex)
                {
                    notes.Add($"style '{style}' could not be applied: {ex.Message}");
                }
            }
            else if (inherited != null)
            {
                try { inserted.set_Style(ref inherited); } catch { }
            }

            try
            {
                if (list == "bullet") inserted.ListFormat.ApplyBulletDefault();
                else if (list == "number") inserted.ListFormat.ApplyNumberDefault();
                else if (newParagraphs) inserted.ListFormat.RemoveNumbers(); // do not continue a list the neighbour is in
            }
            catch (Exception ex)
            {
                if (list == "bullet" || list == "number") notes.Add($"list formatting could not be applied: {ex.Message}");
            }

            var where = newParagraphs || location == "End" ? $"at {location} as new paragraph(s)" : $"at {location}";
            var result = $"Inserted {text?.Length ?? 0} characters {where}";
            return notes.Count > 0 ? result + ", but " + string.Join("; ", notes) + "." : result + ".";
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
