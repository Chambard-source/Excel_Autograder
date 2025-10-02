using ClosedXML.Excel;

public static partial class Grader
{
    private static CheckResult GradeCustomNote(Rule rule, XLWorkbook wbS) 
    {
        var pts = rule.Points;
        var req = rule.Require ?? new RequireSpec();
        bool ok = true;
        var reasons = new List<string>();

        if (req.Sheet is not null && !wbS.Worksheets.Contains(req.Sheet))
        {
            ok = false;
            reasons.Add($"Missing sheet '{req.Sheet}'");
        }

        if (ok && req.PivotTableLike is not null)
        {
            // name-based detection
            bool found =
                wbS.Worksheets.SelectMany(ws => ws.NamedRanges)
                    .Any(nr => (nr.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase))
                || (wbS.NamedRanges != null &&
                    wbS.NamedRanges.Any(nr => (nr.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase)))
                || wbS.Worksheets.SelectMany(ws => ws.Tables)
                    .Any(t => (t.Name ?? "").Contains(req.PivotTableLike, StringComparison.OrdinalIgnoreCase));

            // fallback: visible label on target sheet
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

        return new CheckResult($"custom:{rule.Note ?? "custom"}", pts, ok ? pts : 0, ok,
            reasons.Count == 0 ? "ok" : string.Join("; ", reasons));
    }
}
