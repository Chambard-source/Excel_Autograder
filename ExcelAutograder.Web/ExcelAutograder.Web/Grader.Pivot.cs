using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

public static partial class Grader
{
    private static CheckResult GradePivotLayout(Rule rule, XLWorkbook wbS) 
    {
        var pts = rule.Points;
        var spec = rule.Pivot;
        if (spec is null)
            return new CheckResult("pivot", pts, 0, false, "No pivot spec provided");

        // Helper for section-aware id like "pivot:Summary/SalesPivot"
        string PivotId() =>
            (!string.IsNullOrWhiteSpace(spec.Sheet) || !string.IsNullOrWhiteSpace(spec.TableNameLike))
            ? $"pivot:{spec.Sheet ?? ""}{(!string.IsNullOrWhiteSpace(spec.Sheet) && !string.IsNullOrWhiteSpace(spec.TableNameLike) ? "/" : "")}{spec.TableNameLike ?? ""}"
            : "pivot";

        // choose sheets
        var sheets = string.IsNullOrWhiteSpace(spec.Sheet)
            ? wbS.Worksheets.AsEnumerable()
            : wbS.Worksheets.Where(ws => string.Equals(ws.Name, spec.Sheet, StringComparison.OrdinalIgnoreCase));

        // helpers (reflection-safe)
        static IEnumerable<object> AsEnumerable(object? obj)
            => (obj as System.Collections.IEnumerable)?.Cast<object>() ?? Enumerable.Empty<object>();
        static string S(object? o) => o?.ToString() ?? "";
        static string FirstNonEmpty(params string?[] items)
            => items.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";
        static string? GetStrProp(object o, string name)
            => o.GetType().GetProperty(name)?.GetValue(o)?.ToString();
        static IEnumerable<object> GetEnumProp(object o, string name)
            => AsEnumerable(o.GetType().GetProperty(name)?.GetValue(o));

        static string NormAgg(string raw)
        {
            var a = (raw ?? "").ToLowerInvariant();
            if (a.Contains("sum")) return "sum";
            if (a.Contains("count")) return "count";
            if (a.Contains("avg") || a.Contains("average")) return "average";
            if (a.Contains("min")) return "min";
            if (a.Contains("max")) return "max";
            return string.IsNullOrWhiteSpace(a) ? "sum" : a;
        }

        var findings = new List<string>();

        foreach (var ws in sheets)
        {
            var pivotsObj = ws.GetType().GetProperty("PivotTables")?.GetValue(ws);
            var pivots = AsEnumerable(pivotsObj);
            if (!pivots.Any()) continue;

            foreach (var pt in pivots)
            {
                var ptName = GetStrProp(pt, "Name") ?? "";

                if (!string.IsNullOrWhiteSpace(spec.TableNameLike) &&
                    ptName.IndexOf(spec.TableNameLike, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                HashSet<string> actualRows = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "RowLabels"))
                    actualRows.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));

                HashSet<string> actualCols = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "ColumnLabels"))
                    actualCols.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));

                HashSet<string> actualFilters = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "ReportFilters"))
                    actualFilters.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));

                HashSet<string> actualValues = new(StringComparer.OrdinalIgnoreCase);
                foreach (var v in GetEnumProp(pt, "Values"))
                {
                    var fieldName = FirstNonEmpty(GetStrProp(v, "SourceName"), GetStrProp(v, "CustomName"), GetStrProp(v, "Name"));
                    if (string.IsNullOrWhiteSpace(fieldName)) continue;
                    var sf = GetStrProp(v, "SummaryFormula") ?? GetStrProp(v, "Function") ?? S(v);
                    var agg = NormAgg(sf ?? "");
                    actualValues.Add($"{fieldName}|{agg}");
                }

                var missing = new List<string>();
                if (spec.Rows is { Count: > 0 }) foreach (var need in spec.Rows) if (!actualRows.Contains(need)) missing.Add($"row '{need}'");
                if (spec.Columns is { Count: > 0 }) foreach (var need in spec.Columns) if (!actualCols.Contains(need)) missing.Add($"column '{need}'");
                if (spec.Filters is { Count: > 0 }) foreach (var need in spec.Filters) if (!actualFilters.Contains(need)) missing.Add($"filter '{need}'");
                if (spec.Values is { Count: > 0 })
                    foreach (var need in spec.Values)
                        if (!actualValues.Contains($"{need.Field}|{NormAgg(need.Agg ?? "sum")}"))
                            missing.Add($"value '{need.Field}' with agg '{NormAgg(need.Agg ?? "sum")}'");

                if (missing.Count == 0)
                    return new CheckResult($"pivot:{ws.Name}/{(spec.TableNameLike ?? ptName)}", pts, pts, true,
                        $"pivot '{ptName}' OK");

                findings.Add($"pivot '{ptName}' missing: {string.Join(", ", missing)}");
            }
        }

        if (findings.Count == 0)
            return new CheckResult(PivotId(), pts, 0, false,
                "No pivot tables found or pivot APIs not exposed in this ClosedXML version");

        return new CheckResult(PivotId(), pts, 0, false, string.Join(" | ", findings));
    }
}
