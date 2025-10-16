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
    /// <summary>
    /// Central store for the student's XLSX (zip) bytes for graders that need to inspect
    /// the raw OOXML (charts, conditional formats, pivots). Scoped via <see cref="AsyncLocal{T}"/>.
    /// </summary>
    private static readonly AsyncLocal<byte[]?> _zipBytes = new();

    /// <summary>
    /// Ensures <see cref="_zipBytes"/> is populated with the student's workbook bytes.
    /// Safe to call multiple times; only saves when empty.
    /// </summary>
    /// <param name="wbS">Student workbook.</param>
    private static void EnsureStudentZipBytes(XLWorkbook wbS)
    {
        if (_zipBytes.Value != null) return;
        using var ms = new MemoryStream();
        wbS.SaveAs(ms);
        _zipBytes.Value = ms.ToArray();
    }

    // ---- Entry point used by the web endpoint

    /// <summary>
    /// Grades a student workbook against a rubric using the provided key workbook.
    /// Produces per-check results along with section and sheet metadata for UI grouping.
    /// </summary>
    /// <param name="wbKey">Key (instructor) workbook.</param>
    /// <param name="wbStudent">Student workbook.</param>
    /// <param name="rubric">Rubric definition to apply.</param>
    /// <returns>
    /// An object suitable for the API response containing:
    /// <list type="bullet">
    ///   <item><description><c>score_out_of_total</c> (string)</description></item>
    ///   <item><description><c>score_numeric</c> (double)</description></item>
    ///   <item><description><c>total_points</c> (double)</description></item>
    ///   <item><description><c>details</c> (list of per-check rows with section/sheet/comment/status)</description></item>
    /// </list>
    /// </returns>
    public static object Run(XLWorkbook wbKey, XLWorkbook wbStudent, Rubric rubric)
    {
        // Keep result + section + sheet so the UI can group correctly.
        var results = new List<(CheckResult res, string section, string sheet)>();

        // helper for case-insensitive lookup
        IXLWorksheet? FindWorksheet(XLWorkbook wb, string name) =>
            wb.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));

        foreach (var (sheetName, spec) in rubric.Sheets)
        {
            var wsS = FindWorksheet(wbStudent, sheetName);
            var wsK = FindWorksheet(wbKey, sheetName);

            // If student's sheet is missing, produce one failing row per rule.
            if (wsS is null)
            {
                foreach (var rule in spec.Checks)
                {
                    var sec = string.IsNullOrWhiteSpace(rule.Section) ? "(No section)" : rule.Section.Trim();
                    var name = $"{rule.Type}:{(rule.Cell ?? rule.Range ?? sheetName)}";
                    var msg = $"Sheet '{sheetName}' not found in student workbook.";
                    var res = new CheckResult(name, rule.Points, 0, false, msg);
                    results.Add((res, sec, sheetName));
                }
                continue;
            }

            // ---- compute order
            IEnumerable<Rule> orderedChecks = spec.Checks;

            static string Norm(string? s) => string.IsNullOrWhiteSpace(s) ? "(No section)" : s.Trim();

            // Prefer per-sheet order; else fall back to global meta.sectionOrder
            var order = (spec.SectionOrder is { Count: > 0 })
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
                    .ThenBy(t => t.i)              // stable within section
                    .Select(t => t.r);
            }

            // Grade each rule and attach section/sheet
            foreach (var rule in orderedChecks)
            {
                var sec = string.IsNullOrWhiteSpace(rule.Section) ? "(No section)" : rule.Section.Trim();
                var res = DispatchRule(rule, wbStudent, wbKey, wsS, wsK);

                // If no name was set by the grader, set a sensible default via 'with'
                if (string.IsNullOrWhiteSpace(res.Name))
                {
                    var defaultName = $"{rule.Type}:{(rule.Cell ?? rule.Range ?? "?")}";
                    res = res with { Name = defaultName };
                }

                results.Add((res, sec, sheetName));
            }
        }

        var totalEarned = results.Sum(t => t.res.Earned);
        if (rubric.Scoring?.RoundTo is double roundTo)
            totalEarned = Math.Round(totalEarned, (int)Math.Round(roundTo));

        // Build API rows including section/sheet for correct grouping on the frontend
        var reportRows = results.Select(t =>
        {
            var r = t.res;
            var obj = new Dictionary<string, object?>
            {
                ["check"] = r.Name,
                ["points"] = r.Points,
                ["earned"] = r.Earned,
                ["passed"] = r.Passed,
                ["section"] = t.section,
                ["sheet"] = t.sheet
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

    /// <summary>
    /// Overload that allows the API layer to supply already-saved student XLSX bytes.
    /// This seeds <see cref="_zipBytes"/> for graders that need raw OOXML access.
    /// </summary>
    /// <param name="wbKey">Key workbook.</param>
    /// <param name="wbStudent">Student workbook.</param>
    /// <param name="rubric">Rubric definition.</param>
    /// <param name="studentZipBytes">Raw <c>.xlsx</c> bytes for the student workbook (optional).</param>
    /// <returns>Same response object as <see cref="Run(XLWorkbook, XLWorkbook, Rubric)"/>.</returns>
    public static object Run(XLWorkbook wbKey, XLWorkbook wbStudent, Rubric rubric, byte[]? studentZipBytes)
    {
        _zipBytes.Value = studentZipBytes;
        try { return Run(wbKey, wbStudent, rubric); }
        finally { _zipBytes.Value = null; }
    }

    // ---- Router (kept here for discoverability)

    /// <summary>
    /// Routes a single <see cref="Rule"/> to the appropriate grader based on <c>rule.Type</c>.
    /// </summary>
    /// <param name="rule">The rubric rule to evaluate.</param>
    /// <param name="wbS">Student workbook (some graders need workbook-level context).</param>
    /// <param name="wbK">Key workbook (for key-based comparisons).</param>
    /// <param name="wsS">Student worksheet (for cell/range rules).</param>
    /// <param name="wsK">Key worksheet (optional; null when not present).</param>
    /// <returns>A populated <see cref="CheckResult"/>.</returns>
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
