using System;
using Newtonsoft.Json.Linq;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Executes protocol <c>OfficeOp</c> requests (see packages/protocol) against the
    /// live document through the Office COM object model. Implemented once per
    /// application in the VSTO projects. Always invoked on the Office UI thread.
    /// </summary>
    public interface IOfficeOps
    {
        /// <summary>"Word" | "Excel" | "PowerPoint" (protocol HostApp).</summary>
        string HostApp { get; }

        /// <summary>File name of the active document, or null when nothing is open.</summary>
        string DocumentName { get; }

        /// <summary>Office version/build string for diagnostics.</summary>
        string OfficeVersion { get; }

        /// <summary>Run one operation. Throw <see cref="OfficeOpException"/> for user-facing failures.</summary>
        JToken Execute(JObject op);
    }

    public sealed class OfficeOpException : Exception
    {
        public OfficeOpException(string message) : base(message) { }
        public OfficeOpException(string message, Exception inner) : base(message, inner) { }
    }

    /// <summary>Small helpers shared by the per-application implementations.</summary>
    public static class Ops
    {
        public static string Str(JObject op, string name, string fallback = null)
        {
            var t = op[name];
            return t == null || t.Type == JTokenType.Null ? fallback : t.ToString();
        }

        public static int? Int(JObject op, string name)
        {
            var t = op[name];
            return t == null || t.Type == JTokenType.Null ? (int?)null : t.Value<int>();
        }

        public static bool Bool(JObject op, string name, bool fallback = false)
        {
            var t = op[name];
            return t == null || t.Type == JTokenType.Null ? fallback : t.Value<bool>();
        }

        public static JObject Options(JObject op)
        {
            return op["options"] as JObject ?? new JObject();
        }

        public static string Clip(string text, int max)
        {
            if (text == null) return "";
            return text.Length <= max ? text : text.Substring(0, max) + "…";
        }

        public static JObject Context(string hostApp, string documentName, string fullName, string summary)
        {
            return new JObject
            {
                ["hostApp"] = hostApp,
                ["documentName"] = documentName,
                ["fullName"] = fullName,
                ["summary"] = summary,
            };
        }

        /// <summary>Convert a COM Value2 result (scalar or 2-D array) to a JSON array of rows.</summary>
        public static JArray ToRows(object value)
        {
            var rows = new JArray();
            if (value is object[,] arr)
            {
                var r0 = arr.GetLowerBound(0); var r1 = arr.GetUpperBound(0);
                var c0 = arr.GetLowerBound(1); var c1 = arr.GetUpperBound(1);
                for (var r = r0; r <= r1; r++)
                {
                    var row = new JArray();
                    for (var c = c0; c <= c1; c++) row.Add(ToJson(arr[r, c]));
                    rows.Add(row);
                }
            }
            else
            {
                rows.Add(new JArray(ToJson(value)));
            }
            return rows;
        }

        public static JToken ToJson(object v)
        {
            if (v == null) return JValue.CreateNull();
            switch (v)
            {
                case string s: return s;
                case bool b: return b;
                case double d: return d;
                case int i: return i;
                case DateTime dt: return dt.ToString("o");
                default: return v.ToString();
            }
        }

        /// <summary>Convert a JSON 2-D array to a 1-based object[,] suitable for Range.Value2 / Formula.</summary>
        public static object[,] ToMatrix(JArray rows)
        {
            var rowCount = rows.Count;
            var colCount = 0;
            foreach (var r in rows) colCount = Math.Max(colCount, (r as JArray)?.Count ?? 0);
            var m = new object[rowCount, colCount];
            for (var r = 0; r < rowCount; r++)
            {
                var row = rows[r] as JArray;
                for (var c = 0; c < colCount; c++)
                {
                    var cell = row != null && c < row.Count ? row[c] : null;
                    m[r, c] = FromJson(cell);
                }
            }
            return m;
        }

        public static object FromJson(JToken t)
        {
            if (t == null) return null;
            switch (t.Type)
            {
                case JTokenType.Null: return null;
                case JTokenType.Integer: return t.Value<double>();
                case JTokenType.Float: return t.Value<double>();
                case JTokenType.Boolean: return t.Value<bool>();
                default: return t.ToString();
            }
        }
    }
}
