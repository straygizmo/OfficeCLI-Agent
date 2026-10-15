using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using OfficeCliAgent.Core;
using Office = Microsoft.Office.Core;
using Interop = Microsoft.Office.Interop.PowerPoint;

namespace OfficeCliAgent.PowerPoint
{
    /// <summary>Protocol OfficeOp implementation for PowerPoint (COM).</summary>
    public sealed class PowerPointOps : IOfficeOps
    {
        private readonly Interop.Application _app;

        public PowerPointOps(Interop.Application app) { _app = app; }

        public string HostApp => "PowerPoint";
        public string DocumentName => _app.Presentations.Count > 0 ? _app.ActivePresentation.Name : null;
        public string OfficeVersion => "PowerPoint " + _app.Version + " build " + _app.Build;

        private Interop.Presentation Pres
        {
            get
            {
                if (_app.Presentations.Count == 0) throw new OfficeOpException("No presentation is open in PowerPoint.");
                return _app.ActivePresentation;
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
                case "ppt.setShapeText": return SetShapeText(Ops.Int(op, "slide") ?? 0, Ops.Str(op, "shape"), Ops.Str(op, "text", ""));
                case "ppt.addSlide": return AddSlide(Ops.Int(op, "layout") ?? 2, Ops.Str(op, "title"), Ops.Str(op, "body"), Ops.Int(op, "after"));
                case "ppt.deleteSlides": return DeleteSlides((op["slides"] as JArray)?.Select(t => t.Value<int>()).ToList());
                default: throw new OfficeOpException($"Operation '{kind}' is not supported in PowerPoint.");
            }
        }

        // ------------------------------------------------------------------

        private Interop.Slide Slide(int index)
        {
            var pres = Pres;
            if (index < 1 || index > pres.Slides.Count)
                throw new OfficeOpException($"Slide {index} does not exist (1-{pres.Slides.Count}).");
            return pres.Slides[index];
        }

        private JToken GetContext()
        {
            var pres = Pres;
            var saved = !string.IsNullOrEmpty(pres.Path);
            var sb = new StringBuilder();
            sb.AppendLine("Host: PowerPoint");
            sb.AppendLine("Presentation: " + pres.Name + (saved ? " (" + pres.FullName + ")" : " (unsaved)"));
            sb.AppendLine($"Slides: {pres.Slides.Count}, size: {pres.PageSetup.SlideWidth}x{pres.PageSetup.SlideHeight} pt");
            try { sb.AppendLine("Current slide: " + _app.ActiveWindow.View.Slide.SlideIndex); } catch { /* no window */ }

            var layouts = new List<string>();
            var li = 0;
            foreach (Interop.CustomLayout l in pres.SlideMaster.CustomLayouts) layouts.Add($"{++li}:{l.Name}");
            sb.AppendLine("Layouts: " + string.Join(", ", layouts));

            var limit = Math.Min(pres.Slides.Count, 60);
            for (var i = 1; i <= limit; i++)
            {
                var slide = pres.Slides[i];
                var parts = new List<string>();
                try
                {
                    foreach (Interop.Shape shape in slide.Shapes)
                    {
                        string name;
                        try { name = shape.Name; } catch { name = "(shape)"; }
                        string text = null;
                        try
                        {
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                text = shape.TextFrame.TextRange.Text;
                        }
                        catch { /* charts, OLE objects and SmartArt can throw here */ }
                        string type;
                        try { type = shape.Type.ToString(); } catch { type = "?"; }
                        parts.Add(text != null
                            ? $"{name}: \"{Ops.Clip(text.Replace("\r", " ").Replace("\n", " "), 100)}\""
                            : $"{name} ({type})");
                    }
                }
                catch (Exception ex)
                {
                    parts.Add("(shapes unreadable: " + ex.Message + ")");
                }
                string layoutName;
                try { layoutName = slide.CustomLayout.Name; } catch { layoutName = "?"; }
                sb.AppendLine($"Slide {i} [{layoutName}]: {string.Join(" | ", parts)}");
            }
            if (pres.Slides.Count > limit) sb.AppendLine($"… {pres.Slides.Count - limit} more slides");
            return Ops.Context(HostApp, pres.Name, saved ? pres.FullName : null, sb.ToString());
        }

        private JToken ExportFile(string targetPath)
        {
            if (string.IsNullOrEmpty(targetPath)) throw new OfficeOpException("targetPath is required.");
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
            if (File.Exists(targetPath)) File.Delete(targetPath);
            // Do not embed fonts: unnecessary for a working copy and fragile with restricted (e.g. Japanese) fonts.
            Pres.SaveCopyAs(targetPath, Interop.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            return new JObject { ["path"] = targetPath };
        }

        private JToken ImportFile(string path, JObject options)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) throw new OfficeOpException("File not found: " + path);
            var pres = Pres;
            var index = Ops.Int(options, "afterSlide") ?? pres.Slides.Count;
            if (index < 0 || index > pres.Slides.Count) throw new OfficeOpException($"afterSlide must be between 0 and {pres.Slides.Count}.");
            var from = Ops.Int(options, "fromSlide") ?? 1;
            var to = Ops.Int(options, "toSlide") ?? -1;

            var inserted = pres.Slides.InsertFromFile(path, index, from, to);
            if (Ops.Bool(options, "keepSourceFormatting"))
            {
                // InsertFromFile applies the destination theme. Reapplying the source design per slide
                // is not exposed through COM in a reliable way, so we report it instead of guessing.
                return $"Inserted {inserted} slide(s) from {Path.GetFileName(path)} after slide {index} using the destination theme (keepSourceFormatting is not supported via COM; use officecli on an exported copy if the source design must be preserved).";
            }
            return $"Inserted {inserted} slide(s) from {Path.GetFileName(path)} after slide {index}. Slide count: {pres.Slides.Count}.";
        }

        private JToken SetShapeText(int slideIndex, string shapeName, string text)
        {
            if (string.IsNullOrEmpty(shapeName)) throw new OfficeOpException("shape is required.");
            var slide = Slide(slideIndex);
            Interop.Shape shape;
            try { shape = slide.Shapes[shapeName]; }
            catch { throw new OfficeOpException($"Shape '{shapeName}' not found on slide {slideIndex}. Shapes: {string.Join(", ", slide.Shapes.Cast<Interop.Shape>().Select(s => s.Name))}"); }
            if (shape.HasTextFrame != Office.MsoTriState.msoTrue) throw new OfficeOpException($"Shape '{shapeName}' has no text frame.");
            shape.TextFrame.TextRange.Text = text.Replace("\r\n", "\r").Replace("\n", "\r");
            return $"Set text of '{shapeName}' on slide {slideIndex} ({text.Length} chars).";
        }

        private JToken AddSlide(int layoutIndex, string title, string body, int? after)
        {
            var pres = Pres;
            var layouts = pres.SlideMaster.CustomLayouts;
            if (layoutIndex < 1 || layoutIndex > layouts.Count)
                throw new OfficeOpException($"layout must be between 1 and {layouts.Count}.");
            var position = (after ?? pres.Slides.Count) + 1;
            var slide = pres.Slides.AddSlide(position, layouts[layoutIndex]);

            var filled = new List<string>();
            var placeholders = slide.Shapes.Placeholders;
            if (!string.IsNullOrEmpty(title) && placeholders.Count >= 1)
            {
                placeholders[1].TextFrame.TextRange.Text = title;
                filled.Add("title");
            }
            if (!string.IsNullOrEmpty(body) && placeholders.Count >= 2)
            {
                placeholders[2].TextFrame.TextRange.Text = body.Replace("\r\n", "\r").Replace("\n", "\r");
                filled.Add("body");
            }
            return $"Added slide {slide.SlideIndex} with layout '{layouts[layoutIndex].Name}'" + (filled.Count > 0 ? $" (filled {string.Join(", ", filled)})" : "") + ".";
        }

        private JToken DeleteSlides(List<int> indices)
        {
            if (indices == null || indices.Count == 0) throw new OfficeOpException("slides is required.");
            var pres = Pres;
            foreach (var i in indices.Distinct().OrderByDescending(i => i))
            {
                if (i < 1 || i > pres.Slides.Count) throw new OfficeOpException($"Slide {i} does not exist (1-{pres.Slides.Count}).");
                pres.Slides[i].Delete();
            }
            return $"Deleted {indices.Distinct().Count()} slide(s). Slide count: {pres.Slides.Count}.";
        }
    }
}
