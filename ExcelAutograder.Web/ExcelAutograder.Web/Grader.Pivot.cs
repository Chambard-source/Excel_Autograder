using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Grades a PivotTable’s layout in the student workbook against a <see cref="PivotSpec"/>.
    /// Checks presence of required Rows, Columns, Report Filters, and Values (by caption and
    /// aggregation, with a fallback that infers the source field from captions like
    /// “Sum of Sales”). Supports optional gating by sheet and by table name substring.
    ///
    /// When multiple pivots exist (or name filtering is broad), the method returns success
    /// as soon as one matching pivot is found; otherwise aggregates diagnostics for all
    /// candidate pivots on the inspected sheet(s).
    /// </summary>
    /// <param name="rule">Rule containing <see cref="Rule.Pivot"/> expectations and point value.</param>
    /// <param name="wbS">Student workbook to inspect for pivot tables.</param>
    /// <returns>
    /// <see cref="CheckResult"/> whose <c>Name</c> is <c>pivot:{sheet}/{nameLikeOrFound}</c>.  
    /// Full credit if all required fields match; otherwise 0 with a concise “missing …” summary.  
    /// If no pivots are discoverable (or ClosedXML lacks the needed APIs), returns a failing result with an explanatory comment.
    /// </returns>
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

        static string Norm(string s) => (s ?? "").Trim().ToLowerInvariant();

        static string? TryExtractSourceFromCaption(string caption, string agg)
        {
            var c = (caption ?? "").Trim();
            if (c.Length == 0) return null;

            string[] prefixes = { "sum of ", "average of ", "avg of ", "count of ", "min of ", "max of ", "product of " };
            foreach (var p in prefixes)
                if (c.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                    return c.Substring(p.Length).Trim();

            return null;
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

                HashSet<string> actualCaptionAgg = new(StringComparer.OrdinalIgnoreCase);
                HashSet<string> actualSourceAgg = new(StringComparer.OrdinalIgnoreCase);

                foreach (var v in GetEnumProp(pt, "Values"))
                {
                    var caption = FirstNonEmpty(GetStrProp(v, "CustomName"), GetStrProp(v, "Name"), GetStrProp(v, "SourceName"));
                    var source = GetStrProp(v, "SourceName");
                    var sf = GetStrProp(v, "SummaryFormula") ?? GetStrProp(v, "Function") ?? S(v);
                    var agg = NormAgg(sf ?? "");

                    if (!string.IsNullOrWhiteSpace(caption))
                        actualCaptionAgg.Add($"{caption}|{agg}");
                    if (!string.IsNullOrWhiteSpace(source))
                        actualSourceAgg.Add($"{source}|{agg}");
                }

                var missing = new List<string>();
                if (spec.Rows is { Count: > 0 }) foreach (var need in spec.Rows) if (!actualRows.Contains(need)) missing.Add($"row '{need}'");
                if (spec.Columns is { Count: > 0 }) foreach (var need in spec.Columns) if (!actualCols.Contains(need)) missing.Add($"column '{need}'");
                if (spec.Filters is { Count: > 0 }) foreach (var need in spec.Filters) if (!actualFilters.Contains(need)) missing.Add($"filter '{need}'");
                if (spec.Values is { Count: > 0 })
                {
                    foreach (var need in spec.Values)
                    {
                        var agg = NormAgg(need.Agg ?? "sum");
                        var wantCaption = $"{need.Field}|{agg}";

                        bool ok = actualCaptionAgg.Contains(wantCaption);

                        if (!ok)
                        {
                            var inferredSource = TryExtractSourceFromCaption(need.Field ?? "", agg);
                            if (!string.IsNullOrWhiteSpace(inferredSource))
                                ok = actualSourceAgg.Contains($"{inferredSource}|{agg}");
                        }

                        if (!ok)
                            missing.Add($"value '{need.Field}' with agg '{agg}'");
                    }
                }

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
