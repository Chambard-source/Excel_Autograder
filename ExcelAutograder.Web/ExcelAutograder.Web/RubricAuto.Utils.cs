using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;

public static partial class RubricAuto
{
    /// <summary>
    /// Decide which sheets to include based on the hint and the allSheets flag.
    /// </summary>
    internal static IEnumerable<IXLWorksheet> ResolveSheets(XLWorkbook wb, string? hint, bool all)
    {
        // If user asked for all, return ALL sheets — ignore hint entirely.
        if (all)
            return wb.Worksheets.OrderBy(w => w.Position).ToList();

        // Try explicit hint (exact → contains)
        if (!string.IsNullOrWhiteSpace(hint))
        {
            var hit = wb.Worksheets.FirstOrDefault(w => w.Name.Equals(hint, StringComparison.OrdinalIgnoreCase));
            if (hit != null) return new[] { hit };

            hit = wb.Worksheets.FirstOrDefault(w => w.Name.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0);
            if (hit != null) return new[] { hit };
        }

        // Common fallbacks ("Scores" → "score" → first sheet)
        var scores = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "Scores", StringComparison.OrdinalIgnoreCase));
        if (scores != null) return new[] { scores };

        var score = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "score", StringComparison.OrdinalIgnoreCase));
        if (score != null) return new[] { score };

        return wb.Worksheets.Take(1);
    }

    /// <summary>
    /// Normalize/ensure section order (if you compute it) and perform any per-sheet
    /// cleanup needed right after assembling checks.
    /// </summary>
    internal static void NormalizeOrder(SheetSpec sheet)
    {
        if (sheet?.Checks == null) return;

        foreach (var r in sheet.Checks)
            r.Section = InferSection(r); // keeps explicit section if already set

        // precompute section rank from builder (0,1,2,...) or default to after all
        var indexBySection = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (sheet.SectionOrder != null)
            for (int i = 0; i < sheet.SectionOrder.Count; i++)
                if (!string.IsNullOrWhiteSpace(sheet.SectionOrder[i]))
                    indexBySection[sheet.SectionOrder[i]] = i;

        int SectionRank(Rule r)
            => r.Section != null && indexBySection.TryGetValue(r.Section, out var ix)
               ? ix : int.MaxValue;

        sheet.Checks = sheet.Checks
            .OrderBy(r => SectionRank(r))
            .ThenBy(r => (r.Section ?? "").ToLowerInvariant()) // stable within un-ranked
            .ThenBy(r => TypeRank(r.Type))
            .ThenBy(r => RuleA1Key(r).col)
            .ThenBy(r => RuleA1Key(r).row)
            .ThenBy(r => r.Note ?? "")
            .ToList();
    }

    /// <summary>
    /// Rescales points to the requested total while preserving relative weights.
    /// </summary>
    internal static void RescalePoints(Rubric rub, double targetTotal)
    {
        if (rub is null || rub.Sheets is null || targetTotal <= 0) return;

        var current = 0.0;
        foreach (var ss in rub.Sheets.Values)
            current += ss.Checks?.Sum(c => c.Points) ?? 0.0;

        if (current <= 0) return;

        var scale = targetTotal / current;

        foreach (var ss in rub.Sheets.Values)
            if (ss.Checks is not null)
                foreach (var c in ss.Checks)
                    c.Points = Math.Round(c.Points * scale, 3);

        rub.Points = targetTotal;
    }

    private static bool ContainsAny(string text, params string[] needles)
    => needles.Any(n => text.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0);

    private static bool HasAbsoluteRef(string formula) => formula.IndexOf('$') >= 0;

    private static string NormalizeFormulaAuto(string? f)
    {
        var s = f?.Trim() ?? string.Empty;
        if (s.Length > 0 && s[0] != '=') s = "=" + s;
        return s;
    }

    private static string? KeyCellExpectedText(IXLCell kc)
    {
        // Try Excel’s displayed text first (respects number formats).
        var s = kc.GetFormattedString();
        if (!string.IsNullOrWhiteSpace(s)) return s;

        // If this is a formula, try to read the cached value (ClosedXML doesn’t calculate).
        if (!string.IsNullOrWhiteSpace(kc.FormulaA1))
        {
            // try common property names across CLX versions
            object? cached = kc.GetType().GetProperty("CachedValue")?.GetValue(kc)
                          ?? kc.GetType().GetProperty("ValueCached")?.GetValue(kc);

            if (cached != null)
            {
                // Format cached using the cell’s number format if possible
                try
                {
                    // Temporarily stuff the cached into a clone-like path:
                    // easiest is to format by type without mutating the cell
                    if (cached is double d) return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (cached is bool b) return b ? "TRUE" : "FALSE";
                    if (cached is DateTime dt) return dt.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    return cached.ToString();
                }
                catch { /* fall through */ }
            }
        }

        // Fall back to invariant by data type
        switch (kc.DataType)
        {
            case XLDataType.Number: return kc.GetValue<double>().ToString(CultureInfo.InvariantCulture);
            case XLDataType.Boolean: return kc.GetValue<bool>() ? "TRUE" : "FALSE";
            case XLDataType.DateTime: return kc.GetValue<DateTime>().ToString(CultureInfo.InvariantCulture);
            case XLDataType.Text: return kc.GetString();
            case XLDataType.TimeSpan: return kc.GetValue<TimeSpan>().ToString();
            case XLDataType.Blank: return null;
            default: return kc.Value.ToString();
        }
    }

    private static bool TryToInt(object? v, out int n)
    {
        if (v is int i) { n = i; return true; }
        if (v is double d && Math.Abs(d - Math.Round(d)) < 1e-9)
        { n = (int)Math.Round(d); return true; }
        var s = v?.ToString();
        return int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out n);
    }

    private static bool TryToDouble(object? v, out double d)
    {
        switch (v)
        {
            case null: d = 0; return false;
            case double dx: d = dx; return true;
            case float f: d = f; return true;
            case int i: d = i; return true;
            case long l: d = l; return true;
            case decimal m: d = (double)m; return true;
            case DateTime dt: d = dt.ToOADate(); return true;
            default:
                var s = v.ToString();
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var iv)) { d = iv; return true; }
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out iv)) { d = iv; return true; }
                d = 0; return false;
        }
    }

    // Reflection helpers: -------------------------------------------------------------------------------------------------

    private static IEnumerable<object> AsEnum(object? obj)
        => (obj as System.Collections.IEnumerable)?.Cast<object>() ?? Enumerable.Empty<object>();

    private static string? GetStrProp(object o, string name)
        => o.GetType().GetProperty(name)?.GetValue(o)?.ToString();

    private static IEnumerable<object> GetEnumProp(object o, string name)
        => AsEnum(o.GetType().GetProperty(name)?.GetValue(o));

    private static string FirstNonEmpty(params string?[] xs)
        => xs.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";

    private static string NormAgg(string? raw)
    {
        var a = (raw ?? "").ToLowerInvariant();
        if (a.Contains("sum")) return "sum";
        if (a.Contains("count")) return "count";
        if (a.Contains("avg") || a.Contains("average")) return "average";
        if (a.Contains("min")) return "min";
        if (a.Contains("max")) return "max";
        return string.IsNullOrWhiteSpace(a) ? "sum" : a;
    }

    // Safely rescales every rule's points so the rubric totals to desiredTotal.


    // Normalize any incoming address-ish string:
    //   - takes "Sheet Name!$E$8" → "E8"
    //   - takes "$E$8" → "E8"
    //   - takes "E7:E18" → "E7"
    //   - leaves "E8" as-is
    private static string CleanA1(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var t = s.Trim();

        // strip sheet prefix
        var bang = t.LastIndexOf('!');
        if (bang >= 0 && bang < t.Length - 1) t = t.Substring(bang + 1);

        // if it's a range, keep the first corner
        var colon = t.IndexOf(':');
        if (colon > 0) t = t.Substring(0, colon);

        // remove $ anchors
        t = t.Replace("$", "");

        return t;
    }

    // Converts A1 to sortable (col,row); non-address → (Max,Max)
    private static (int col, int row) A1ToKey(string? a1Raw)
    {
        var a1 = CleanA1(a1Raw);
        if (string.IsNullOrEmpty(a1)) return (int.MaxValue, int.MaxValue);

        int i = 0;
        while (i < a1.Length && char.IsLetter(a1[i])) i++;

        if (i == 0) return (int.MaxValue, int.MaxValue);   // no letters

        var colStr = a1.Substring(0, i).ToUpperInvariant();
        var rowStr = a1.Substring(i);

        int col = 0;
        foreach (var ch in colStr) col = col * 26 + (ch - 'A' + 1);

        if (!int.TryParse(rowStr, out var row)) row = int.MaxValue;
        return (col, row);
    }

    // Helper: best A1 for a rule (cell first, else range’s first corner)
    private static (int col, int row) RuleA1Key(Rule r)
    {
        if (!string.IsNullOrWhiteSpace(r.Cell)) return A1ToKey(r.Cell);
        if (!string.IsNullOrWhiteSpace(r.Range)) return A1ToKey(r.Range);
        return (int.MaxValue, int.MaxValue);
    }

    // Give common rule types a stable type-rank
    private static int TypeRank(string? t) => t?.ToLowerInvariant() switch
    {
        "format" => 0,
        "table" => 1,
        "pivot" => 2,
        "formula" => 3,
        "chart" => 4,
        "custom_note" => 9,
        _ => 5
    };

    // Best-effort section inference so things cluster automatically
    private static string InferSection(Rule r)
    {
        // honor explicit section if author provided one
        if (!string.IsNullOrWhiteSpace(r.Section)) return r.Section!;

        string t = (r.Type ?? "").ToLowerInvariant();
        string note = (r.Note ?? "").ToLowerInvariant();
        string cf = (r.ExpectedFormula ?? r.ExpectedFormulaRegex ?? r.Expected?.ToString() ?? "").ToLowerInvariant();
        string cell = (r.Cell ?? "").ToUpperInvariant();

        // broad buckets
        if (t == "chart") return "Charts";
        if (t == "table") return "Tables";
        if (t == "format") return "Header / Formatting";
        if (note.Contains("header")) return "Header / Formatting";
        if (note.Contains("pivot") || cf.Contains("pivot")) return "Pivot";

        if (t == "formula")
        {
            // light Excel-HW heuristics
            if (cf.Contains("sumif(")) return "SUMIF / category totals";
            if (cf.StartsWith("=$e$") || cf.Contains("$e$27")) return "Relative frequency";
            if (cell.StartsWith("G") && !cf.Contains("count")) return "Percent formatting";
            if (cf.StartsWith("sum(") || cf.Contains(":E17)") || cf.Contains(":L18)")) return "Totals";
        }

        return "Other";
    }

    private static string? CleanExpected(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return s;

        // Convert NBSP to space, drop control chars, collapse & trim
        var cleaned = new string(s
            .Select(ch => ch == '\u00A0' ? ' ' : ch) // nbsp -> space
            .Where(ch => !char.IsControl(ch) || ch == '\n' || ch == '\r' || ch == '\t')
            .ToArray());

        // normalize whitespace: newlines/tabs -> space, collapse, trim
        cleaned = cleaned.Replace('\n', ' ').Replace('\r', ' ').Replace('\t', ' ');
        while (cleaned.Contains("  ")) cleaned = cleaned.Replace("  ", " ");
        return cleaned.Trim();
    }
}
