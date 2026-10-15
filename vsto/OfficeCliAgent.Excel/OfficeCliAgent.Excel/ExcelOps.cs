using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using OfficeCliAgent.Core;
using Interop = Microsoft.Office.Interop.Excel;

namespace OfficeCliAgent.Excel
{
    /// <summary>Protocol OfficeOp implementation for Excel (COM).</summary>
    public sealed class ExcelOps : IOfficeOps
    {
        private const int PreviewRows = 40;
        private const int MaxCells = 20000;
        private readonly Interop.Application _app;

        public ExcelOps(Interop.Application app) { _app = app; }

        public string HostApp => "Excel";
        public string DocumentName => _app.Workbooks.Count > 0 ? _app.ActiveWorkbook.Name : null;
        public string OfficeVersion => "Excel " + _app.Version + " build " + _app.Build;

        private Interop.Workbook Book
        {
            get
            {
                if (_app.Workbooks.Count == 0) throw new OfficeOpException("No workbook is open in Excel.");
                return _app.ActiveWorkbook;
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
                case "excel.listSheets": return ListSheets();
                case "excel.getRange": return GetRange(Ops.Str(op, "sheet"), Ops.Str(op, "address"));
                case "excel.setRange": return SetRange(Ops.Str(op, "sheet"), Ops.Str(op, "address"), op["values"] as JArray, op["formulas"] as JArray, Ops.Str(op, "numberFormat"));
                default: throw new OfficeOpException($"Operation '{kind}' is not supported in Excel.");
            }
        }

        // ------------------------------------------------------------------

        private Interop.Worksheet Sheet(string name)
        {
            var book = Book;
            if (string.IsNullOrEmpty(name)) return (Interop.Worksheet)book.ActiveSheet;
            foreach (Interop.Worksheet ws in book.Worksheets)
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
            throw new OfficeOpException($"Worksheet '{name}' not found. Sheets: {string.Join(", ", SheetNames(book))}");
        }

        private static IEnumerable<string> SheetNames(Interop.Workbook book)
        {
            foreach (Interop.Worksheet ws in book.Worksheets) yield return ws.Name;
        }

        private JToken GetContext()
        {
            var book = Book;
            var saved = !string.IsNullOrEmpty(book.Path);
            var sb = new StringBuilder();
            sb.AppendLine("Host: Excel");
            sb.AppendLine("Workbook: " + book.Name + (saved ? " (" + book.FullName + ")" : " (unsaved)"));
            sb.AppendLine("Sheets: " + string.Join(", ", SheetNames(book).Select(n => n)));
            var active = (Interop.Worksheet)book.ActiveSheet;
            sb.AppendLine("Active sheet: " + active.Name);
            try
            {
                if (_app.Selection is Interop.Range sel) sb.AppendLine("Selection: " + sel.Address[false, false]);
            }
            catch { /* selection may not be a range */ }

            var used = active.UsedRange;
            if (used == null || used.Count == 0 || (used.Rows.Count == 1 && used.Columns.Count == 1 && used.Value2 == null))
            {
                sb.AppendLine("Used range: (empty)");
            }
            else
            {
                sb.AppendLine($"Used range: {used.Address[false, false]} ({used.Rows.Count} x {used.Columns.Count})");
                var rows = Math.Min(PreviewRows, used.Rows.Count);
                var cols = Math.Min(50, used.Columns.Count);
                var preview = (Interop.Range)used.Resize[rows, cols];
                var values = Ops.ToRows(preview.Value2);
                sb.AppendLine("First rows:");
                foreach (JArray row in values)
                    sb.AppendLine(string.Join("\t", row.Select(c => Ops.Clip(c.Type == JTokenType.Null ? "" : c.ToString(), 30))));
            }
            return Ops.Context(HostApp, book.Name, saved ? book.FullName : null, sb.ToString());
        }

        private JToken ExportFile(string targetPath)
        {
            if (string.IsNullOrEmpty(targetPath)) throw new OfficeOpException("targetPath is required.");
            Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
            if (File.Exists(targetPath)) File.Delete(targetPath);
            Book.SaveCopyAs(targetPath);
            return new JObject { ["path"] = targetPath };
        }

        private JToken ImportFile(string path, JObject options)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) throw new OfficeOpException("File not found: " + path);
            var wanted = (options["sheetNames"] as JArray)?.Select(t => t.ToString()).ToList();
            var afterName = Ops.Str(options, "afterSheet");
            var book = Book;

            Interop.Worksheet after = null;
            if (!string.IsNullOrEmpty(afterName)) after = Sheet(afterName);
            else after = (Interop.Worksheet)book.Worksheets[book.Worksheets.Count];

            var screen = _app.ScreenUpdating;
            _app.ScreenUpdating = false;
            Interop.Workbook src = null;
            var copied = new List<string>();
            try
            {
                src = _app.Workbooks.Open(path, ReadOnly: true, AddToMru: false, UpdateLinks: 0);
                foreach (Interop.Worksheet ws in src.Worksheets)
                {
                    if (wanted != null && !wanted.Any(w => string.Equals(w, ws.Name, StringComparison.OrdinalIgnoreCase))) continue;
                    ws.Copy(After: after);
                    after = (Interop.Worksheet)book.Worksheets[after.Index + 1];
                    copied.Add(after.Name);
                }
                if (wanted != null && copied.Count == 0)
                    throw new OfficeOpException($"None of the requested sheets exist in {Path.GetFileName(path)}. Available: {string.Join(", ", SheetNames(src))}");
            }
            finally
            {
                if (src != null) src.Close(SaveChanges: false);
                _app.ScreenUpdating = screen;
                book.Activate();
            }
            return $"Inserted {copied.Count} sheet(s) from {Path.GetFileName(path)}: {string.Join(", ", copied)}";
        }

        private JToken ListSheets()
        {
            var arr = new JArray();
            foreach (Interop.Worksheet ws in Book.Worksheets)
            {
                var used = ws.UsedRange;
                arr.Add(new JObject
                {
                    ["name"] = ws.Name,
                    ["visible"] = ws.Visible == Interop.XlSheetVisibility.xlSheetVisible,
                    ["usedRange"] = used?.Address[false, false],
                });
            }
            return arr;
        }

        private JToken GetRange(string sheetName, string address)
        {
            var ws = Sheet(sheetName);
            var range = string.IsNullOrEmpty(address) ? ws.UsedRange : ws.Range[address];
            if (range.Count > MaxCells)
                throw new OfficeOpException($"Range {range.Address[false, false]} has {range.Count} cells; limit is {MaxCells}. Read a smaller range or export the workbook and use officecli.");
            return new JObject
            {
                ["sheet"] = ws.Name,
                ["address"] = range.Address[false, false],
                ["values"] = Ops.ToRows(range.Value2),
                ["formulas"] = Ops.ToRows(range.Formula),
            };
        }

        private JToken SetRange(string sheetName, string address, JArray values, JArray formulas, string numberFormat)
        {
            if (string.IsNullOrEmpty(address)) throw new OfficeOpException("address is required.");
            var ws = Sheet(sheetName);
            var range = ws.Range[address];
            var written = 0;

            if (values != null)
            {
                var m = Ops.ToMatrix(values);
                var target = (Interop.Range)range.Resize[m.GetLength(0), m.GetLength(1)];
                target.Value2 = m;
                written += m.Length;
            }
            if (formulas != null)
            {
                var m = Ops.ToMatrix(formulas);
                var target = (Interop.Range)range.Resize[m.GetLength(0), m.GetLength(1)];
                target.Formula = m;
                written += m.Length;
            }
            if (!string.IsNullOrEmpty(numberFormat)) range.NumberFormat = numberFormat;

            return $"Wrote {written} cell(s) at {ws.Name}!{range.Address[false, false]}" + (numberFormat != null ? $" (format {numberFormat})" : "") + ".";
        }
    }
}
