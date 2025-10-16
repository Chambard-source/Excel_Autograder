using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Grades a custom, note-based requirement that may reference a target sheet and/or a "pivot-like" artifact.
    /// </summary>
    /// <param name="rule">
    /// The rubric rule. Uses <c>rule.Points</c>, <c>rule.Note</c>, and (optionally) <c>rule.Require</c>:
    /// <list type="bullet">
    ///   <item><description><c>Require.Sheet</c> (string): sheet that must exist</description></item>
    ///   <item><description><c>Require.PivotTableLike</c> (string): substring to match against pivot/table/named-range names or visible cell text</description></item>
    /// </list>
    /// </param>
    /// <param name="wbS">The student's workbook being graded.</param>
    /// <returns>
    /// A <see cref="CheckResult"/> whose id is <c>"custom:{rule.Note}"</c>, awarding all points if found, else 0.
    /// </returns>
    /// <remarks>
    /// Detection heuristic when <c>Require.PivotTableLike</c> is provided:
    /// <list type="number">
    ///   <item><description>Search all worksheet-scoped named ranges</description></item>
    ///   <item><description>Search workbook-scoped named ranges</description></item>
    ///   <item><description>Search all Excel tables on all sheets</description></item>
    ///   <item><description>If a target sheet is specified and still not found, scan visible cell text on that sheet</description></item>
    /// </list>
    /// If <c>Require.Sheet</c> is set and missing, grading fails immediately.
    /// </remarks>
    private static CheckResult GradeCustomNote(Rule rule, XLWorkbook wbS)
    {
        var pts = rule.Points;
        var req = rule.Require ?? new RequireSpec();
        bool ok = true;
        var reasons = new List<string>();

        // 1) Sheet existence (hard requirement if provided)
        if (req.Sheet is not null && !wbS.Worksheets.Contains(req.Sheet))
        {
            ok = false;
            reasons.Add($"Missing sheet '{req.Sheet}'");
        }

        // 2) "Pivot-like" detection across names/tables and (optionally) visible text
        if (ok && req.PivotTableLike is not null)
        {
            // Name-based detection across all sheets and workbook scope
            bool found =
                wbS.Worksheets.SelectMany(ws => ws.NamedRanges)
                    .Any(nr => (nr.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase))
                || (wbS.NamedRanges != null &&
                    wbS.NamedRanges.Any(nr => (nr.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase)))
                || wbS.Worksheets.SelectMany(ws => ws.Tables)
                    .Any(t => (t.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase));

            // Fallback: scan visible text on the specified sheet only (cheaper than all sheets)
            if (!found && req.Sheet is not null && wbS.Worksheets.Contains(req.Sheet))
            {
                var ws = wbS.Worksheets.Worksheet(req.Sheet);
                foreach (var c in ws.CellsUsed(XLCellsUsedOptions.All))
                {
                    var text = c.GetString();
                    if (!string.IsNullOrEmpty(text) &&
                        text.Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase))
                    {
                        found = true;
                        break;
                    }
                }
            }

            if (!found)
            {
                ok = false;
                reasons.Add($"Pivot-like '{req.PivotTableLike}' not found");
            }
        }

        return new CheckResult(
            $"custom:{rule.Note ?? "custom"}",
            pts,
            ok ? pts : 0,
            ok,
            reasons.Count == 0 ? "ok" : string.Join("; ", reasons)
        );
    }
}
