using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    // Central store for the student's XLSX bytes (used by chart/CF/pivot graders)
    private static readonly AsyncLocal<byte[]?> _zipBytes = new();

    private static void EnsureStudentZipBytes(XLWorkbook wbS)
    {
        if (_zipBytes.Value != null) return;
        using var ms = new MemoryStream();
        wbS.SaveAs(ms);
        _zipBytes.Value = ms.ToArray();
    }

    // ---- Entry point used by the web endpoint
    public static object Run(XLWorkbook wbKey, XLWorkbook wbStudent, Rubric rubric)
    {
        var results = new List<CheckResult>();

        // helper for case-insensitive lookup
        IXLWorksheet? FindWorksheet(XLWorkbook wb, string name) =>
            wb.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));

        foreach (var (sheetName, spec) in rubric.Sheets)
        {
            var wsS = FindWorksheet(wbStudent, sheetName);
            if (wsS is null)
            {
                foreach (var rule in spec.Checks)
                {
                    var id = rule.Cell ?? rule.Range ?? sheetName;
                    results.Add(new CheckResult($"{rule.Type}:{id}", rule.Points, 0, false,
                        $"Sheet '{sheetName}' missing"));
                }
                continue;
            }

            var wsK = FindWorksheet(wbKey, sheetName);

            // ---- compute order
            IEnumerable<Rule> orderedChecks = spec.Checks;

            // normalize names used as keys
            string Norm(string? s) => string.IsNullOrWhiteSpace(s) ? "(No section)" : s.Trim();

            // Prefer per-sheet; else fall back to global meta.sectionOrder
            var order = spec.SectionOrder?.Count > 0
                ? spec.SectionOrder
                : rubric.Meta?.SectionOrder;

            if (order is { Count: > 0 })
            {
                var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < order.Count; i++) index[Norm(order[i])] = i;

                int Rank(Rule r) => index.TryGetValue(Norm(r.Section), out var pos) ? pos : int.MaxValue;

                orderedChecks = spec.Checks
                    .Select((r, i) => (r, i))
                    .OrderBy(t => Rank(t.r))       // section order
                    .ThenBy(t => t.i)              // stable within a section
                    .Select(t => t.r);
            }

            foreach (var rule in orderedChecks)
                results.Add(DispatchRule(rule, wbStudent, wbKey, wsS, wsK));
        }

        var totalEarned = results.Sum(r => r.Earned);
        if (rubric.Scoring?.RoundTo is double roundTo)
            totalEarned = Math.Round(totalEarned, (int)Math.Round(roundTo));

        var reportRows = results.Select(r =>
        {
            var obj = new Dictionary<string, object?>
            {
                ["check"] = r.Name,
                ["points"] = r.Points,
                ["earned"] = r.Earned,
                ["passed"] = r.Passed
            };
            if (rubric.Report?.IncludeComments != false) obj["comment"] = r.Comment;
            if (rubric.Report?.IncludePassFailColumn != false) obj["status"] = r.Passed ? "PASS" : "FAIL";
            return obj;
        }).ToList();

        return new Dictionary<string, object?>
        {
            ["score_out_of_total"] = $"{totalEarned}/{rubric.Points}",
            ["score_numeric"] = totalEarned,
            ["total_points"] = rubric.Points,
            ["details"] = reportRows
        };
    }

    // Variant that lets the API pass pre-saved student xlsx bytes
    public static object Run(XLWorkbook wbKey, XLWorkbook wbStudent, Rubric rubric, byte[]? studentZipBytes)
    {
        _zipBytes.Value = studentZipBytes;
        try { return Run(wbKey, wbStudent, rubric); }
        finally { _zipBytes.Value = null; }
    }

    // ---- Router (kept here for discoverability)
    private static CheckResult DispatchRule(Rule rule, XLWorkbook wbS, XLWorkbook wbK, IXLWorksheet wsS, IXLWorksheet? wsK)
        => rule.Type.ToLowerInvariant() switch
        {
            "value" => GradeValue(rule, wsS, wsK),
            "formula" => GradeFormula(rule, wsS, wsK),
            "format" => GradeFormat(rule, wsS),
            "range_value" => GradeRangeValue(rule, wsS, wsK),
            "range_formula" => GradeRangeFormula(rule, wsS, wsK),
            "range_format" => GradeRangeFormat(rule, wsS),
            "custom_note" => GradeCustomNote(rule, wbS),
            "range_sequence" => GradeRangeSequence(rule, wsS),
            "range_numeric" => GradeRangeNumeric(rule, wsS),
            "chart" => GradeChart(rule, wbS),
            "pivot_layout" => GradePivotLayout(rule, wbS),
            "conditional_format" => GradeConditionalFormat(rule, wbS),
            "table" => GradeTable(rule, wsS),
            _ => new CheckResult(rule.Type, rule.Points, 0, false, $"Unknown check type '{rule.Type}'")
        };
}
