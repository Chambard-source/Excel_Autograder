using System.Globalization;
using ClosedXML.Excel;

public static partial class RubricAuto
{
    public static Rubric GenerateFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets)
        => BuildFromKey(wbKey, sheetHint, allSheets, 5.0, keyZipBytes: null);

    public static void ScalePoints(Rubric rub, double target) => RescalePoints(rub, target);

    public static Rubric BuildFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets, double targetTotal)
        => BuildFromKey(wbKey, sheetHint, allSheets, targetTotal, keyZipBytes: null);

    public static Rubric BuildFromKey(
        XLWorkbook wbKey, string? sheetHint, bool allSheets, double targetTotal, byte[]? keyZipBytes)
    {
        var zipPivots = keyZipBytes is null ? new(StringComparer.OrdinalIgnoreCase) : ExtractPivotRulesFromZip(keyZipBytes);
        var zipCF = keyZipBytes is null ? new(StringComparer.OrdinalIgnoreCase) : ExtractConditionalRulesFromZip(keyZipBytes);
        var zipCharts = keyZipBytes is null ? new(StringComparer.OrdinalIgnoreCase) : ExtractChartRulesFromZip(keyZipBytes);

        var rub = new Rubric { Points = 0, Sheets = new(StringComparer.OrdinalIgnoreCase), Report = new Report { IncludePassFailColumn = true, IncludeComments = true } };

        var sheets = ResolveSheets(wbKey, sheetHint, allSheets);
        foreach (var ws in sheets)
        {
            var checks = BuildChecksForSheet(ws);

            if (zipPivots.TryGetValue(ws.Name, out var pr) && pr.Count > 0)
                MergePivotRules(checks, pr);

            if (zipCF.TryGetValue(ws.Name, out var cr) && cr.Count > 0)
                checks.AddRange(cr);

            if (zipCharts.TryGetValue(ws.Name, out var zr) && zr.Count > 0)
                checks.AddRange(zr);

            if (checks.Count == 0)
                checks.Add(new Rule { Type = "custom_note", Points = 0, Note = "No auto-generated checks for this sheet" });

            rub.Sheets[ws.Name] = new SheetSpec { Checks = checks };
            NormalizeOrder(rub.Sheets[ws.Name]);
            rub.Points += checks.Sum(c => c.Points);
        }

        foreach (var kv in zipCharts)
            if (!rub.Sheets.ContainsKey(kv.Key))
            {
                rub.Sheets[kv.Key] = new SheetSpec { Checks = kv.Value };
                NormalizeOrder(rub.Sheets[kv.Key]);
                rub.Points += kv.Value.Sum(r => r.Points);
            }

        if (targetTotal > 0) RescalePoints(rub, targetTotal);
        return rub;
    }

    private static void MergePivotRules(List<Rule> checks, List<Rule> zipRules)
    {
        var existing = new HashSet<string>(
            checks.Where(c => c.Type == "pivot_layout" && c.Pivot?.TableNameLike != null)
                  .Select(c => c.Pivot!.TableNameLike!), StringComparer.OrdinalIgnoreCase);

        foreach (var r in zipRules)
            if (r.Pivot?.TableNameLike is null || !existing.Contains(r.Pivot.TableNameLike))
                checks.Add(r);
    }

    // Build rubric from explicit sections & ranges per sheet.
    // sectionsPerSheet: Sheet -> list of (sectionName, ranges[])
    public static Rubric BuildFromKeyRanges(
        XLWorkbook wbKey,
        IDictionary<string, List<(string section, List<string> ranges)>> sectionsPerSheet,
        bool includeArtifacts,
        double targetTotal,
        byte[]? keyZipBytes)
    {
        var rub = new Rubric
        {
            Points = 0,
            Sheets = new Dictionary<string, SheetSpec>(StringComparer.OrdinalIgnoreCase),
            Report = new Report { IncludePassFailColumn = true, IncludeComments = true }
        };

        // Optional artifact discovery via ZIP (reuse your existing extractors)
        var pivots = includeArtifacts && keyZipBytes != null ? ExtractPivotRulesFromZip(keyZipBytes) : new();
        var cfs = includeArtifacts && keyZipBytes != null ? ExtractConditionalRulesFromZip(keyZipBytes) : new();
        var charts = includeArtifacts && keyZipBytes != null ? ExtractChartRulesFromZip(keyZipBytes) : new();

        foreach (var (sheetName, sections) in sectionsPerSheet)
        {
            var ws = wbKey.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (ws is null) continue;

            var checks = new List<Rule>();
            var sectionOrder = new List<string>();

            foreach (var (sectionName, ranges) in sections)
            {
                // remember order
                if (!string.IsNullOrWhiteSpace(sectionName)) sectionOrder.Add(sectionName);

                // parse ranges on this sheet
                var addrs = new List<IXLRangeAddress>();
                foreach (var r in ranges ?? Enumerable.Empty<string>())
                {
                    try { addrs.Add(ws.Range(r).RangeAddress); } catch { /* skip malformed */ }
                }
                if (addrs.Count == 0) continue;

                // 1) Formula cells inside the ranges
                foreach (var ra in addrs)
                {
                    foreach (var cell in ws.Range(ra).CellsUsed(XLCellsUsedOptions.All))
                    {
                        var f = NormalizeFormulaAuto(cell.FormulaA1);
                        if (string.IsNullOrWhiteSpace(f)) continue;

                        var rule = new Rule
                        {
                            Type = "formula",
                            Points = 0.5,
                            Cell = cell.Address.ToString(),
                            Note = "Formula from key (selected range)",
                            ExpectedFromKey = true,
                            ExpectedFormula = f,
                            Expected = System.Text.Json.JsonSerializer.SerializeToElement(CleanExpected(KeyCellExpectedText(cell))),
                            Section = sectionName
                        };
                        if (HasAbsoluteRef(f)) rule.RequireAbsolute = true;
                        checks.Add(rule);
                    }
                }

                // 2) Per-cell VALUE rules for non-formula cells in each range
                foreach (var ra in addrs)
                {
                    var rangeObj = ws.Range(ra);
                    foreach (var cell in rangeObj.CellsUsed(XLCellsUsedOptions.All))
                    {
                        // Skip cells that already produced a formula rule
                        if (!string.IsNullOrWhiteSpace(cell.FormulaA1)) continue;

                        var a1 = cell.Address.ToStringRelative();

                        checks.Add(new Rule
                        {
                            Type = "value",
                            Points = 0.5, // tune as you like; rescaled later
                            Cell = a1,
                            Note = "Value from key (selected range)",
                            ExpectedFromKey = true,
                            Expected = System.Text.Json.JsonSerializer.SerializeToElement(CleanExpected(KeyCellExpectedText(cell))),
                            Section = sectionName
                        });
                    }
                }

                // 3) Summary row immediately below each range
                foreach (var ra in addrs)
                {
                    var rr = ws.Range(ra);
                    var lastRow = rr.RangeAddress.LastAddress.RowNumber;
                    var firstCol = rr.RangeAddress.FirstAddress.ColumnNumber;
                    var lastCol = rr.RangeAddress.LastAddress.ColumnNumber;

                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        var cell = ws.Cell(lastRow + 1, col);      // public API
                        var f = NormalizeFormulaAuto(cell.FormulaA1);
                        if (string.IsNullOrWhiteSpace(f)) continue;

                        var rule = new Rule
                        {
                            Type = "formula",
                            Points = 1.0,
                            Cell = cell.Address.ToString(),
                            Note = "Summary formula (selected range)",
                            ExpectedFromKey = true,
                            ExpectedFormula = f,
                            Expected = System.Text.Json.JsonSerializer.SerializeToElement(CleanExpected(KeyCellExpectedText(cell))),
                            Section = sectionName
                        };
                        if (HasAbsoluteRef(f)) rule.RequireAbsolute = true;
                        checks.Add(rule);
                    }
                }
            }

            // Add artifacts (optional) under "Artifacts" section
            if (includeArtifacts)
            {
                bool anyArtifacts = false;
                if (pivots.TryGetValue(ws.Name, out var pr)) { foreach (var r in pr) { r.Section ??= "Artifacts"; checks.Add(r); } anyArtifacts |= pr.Count > 0; }
                if (cfs.TryGetValue(ws.Name, out var cr)) { foreach (var r in cr) { r.Section ??= "Artifacts"; checks.Add(r); } anyArtifacts |= cr.Count > 0; }
                if (charts.TryGetValue(ws.Name, out var ch)) { foreach (var r in ch) { r.Section ??= "Artifacts"; checks.Add(r); } anyArtifacts |= ch.Count > 0; }
                if (anyArtifacts && !sectionOrder.Contains("Artifacts")) sectionOrder.Add("Artifacts");
            }

            var spec = new SheetSpec { Checks = checks, SectionOrder = sectionOrder };
            NormalizeOrder(spec);
            rub.Sheets[ws.Name] = spec;
            rub.Points += checks.Sum(c => c.Points);
        }

        if (targetTotal > 0) RescalePoints(rub, targetTotal);
        return rub;
    }

    public static Rubric BuildFromKeyRanges(
    XLWorkbook wbKey,
    IDictionary<string, List<string>> ranges,
    bool includeArtifacts,
    double targetTotal,
    byte[]? keyZipBytes)
    {
        var rub = new Rubric
        {
            Points = 0,
            Sheets = new Dictionary<string, SheetSpec>(StringComparer.OrdinalIgnoreCase),
            Report = new Report { IncludePassFailColumn = true, IncludeComments = true }
        };

        var zipPivots = includeArtifacts && keyZipBytes != null ? ExtractPivotRulesFromZip(keyZipBytes) : new();
        var zipCF = includeArtifacts && keyZipBytes != null ? ExtractConditionalRulesFromZip(keyZipBytes) : new();
        var zipCharts = includeArtifacts && keyZipBytes != null ? ExtractChartRulesFromZip(keyZipBytes) : new();

        foreach (var kv in ranges)
        {
            var ws = wbKey.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, kv.Key, StringComparison.OrdinalIgnoreCase));
            if (ws is null) continue;

            var addrs = new List<IXLRangeAddress>();
            foreach (var r in kv.Value ?? Enumerable.Empty<string>())
            {
                try { addrs.Add(ws.Range(r).RangeAddress); } catch { }
            }
            if (addrs.Count == 0) continue;

            var checks = BuildChecksForRanges(ws, addrs);

            if (includeArtifacts)
            {
                if (zipPivots.TryGetValue(ws.Name, out var pr)) checks.AddRange(pr);
                if (zipCF.TryGetValue(ws.Name, out var cr)) checks.AddRange(cr);
                if (zipCharts.TryGetValue(ws.Name, out var ch)) checks.AddRange(ch);
            }

            if (checks.Count == 0)
                checks.Add(new Rule { Type = "custom_note", Points = 0, Note = "No auto-generated checks in the selected ranges" });

            rub.Sheets[ws.Name] = new SheetSpec { Checks = checks };
            NormalizeOrder(rub.Sheets[ws.Name]);
            rub.Points += checks.Sum(c => c.Points);
        }

        if (targetTotal > 0) RescalePoints(rub, targetTotal);
        return rub;
    }
}
