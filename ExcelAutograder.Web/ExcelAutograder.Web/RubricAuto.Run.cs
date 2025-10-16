using System.Globalization;
using ClosedXML.Excel;

public static partial class RubricAuto
{
    /// <summary>
    /// Convenience entry point that builds a rubric from a key workbook using default total points (5.0).
    /// Persists a ZIP snapshot of the workbook so artifact extractors (pivots/CF/charts) can run.
    /// </summary>
    /// <param name="wbKey">The key (instructor) workbook.</param>
    /// <param name="sheetHint">Optional sheet name or substring hint to restrict which sheets are analyzed.</param>
    /// <param name="allSheets">If true, analyze all sheets and ignore <paramref name="sheetHint"/>.</param>
    /// <returns>A constructed <see cref="Rubric"/> with auto-generated checks and scaled totals.</returns>
    public static Rubric GenerateFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets)
    {
        using var ms = new MemoryStream();
        wbKey.SaveAs(ms);
        return BuildFromKey(wbKey, sheetHint, allSheets, 5.0, keyZipBytes: ms.ToArray());
    }

    /// <summary>
    /// Rescales every rule’s points in-place so the rubric total equals <paramref name="target"/>.
    /// Relative weights are preserved.
    /// </summary>
    /// <param name="rub">Rubric to modify.</param>
    /// <param name="target">Desired total points for the rubric.</param>
    public static void ScalePoints(Rubric rub, double target) => RescalePoints(rub, target);

    /// <summary>
    /// Builds a rubric from the key workbook and scales to <paramref name="targetTotal"/>.
    /// Creates a ZIP snapshot for artifact extraction.
    /// </summary>
    /// <param name="wbKey">The key workbook.</param>
    /// <param name="sheetHint">Optional sheet name or substring hint.</param>
    /// <param name="allSheets">Set to true to include all sheets.</param>
    /// <param name="targetTotal">Desired total points for the rubric.</param>
    /// <returns>A constructed, scaled <see cref="Rubric"/>.</returns>
    public static Rubric BuildFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets, double targetTotal)
    {
        using var ms = new MemoryStream();
        wbKey.SaveAs(ms);
        return BuildFromKey(wbKey, sheetHint, allSheets, targetTotal, keyZipBytes: ms.ToArray());
    }

    /// <summary>
    /// Core builder that constructs a rubric from a key workbook.
    /// Optionally consumes a pre-saved workbook ZIP to extract artifacts like pivot layouts,
    /// conditional formats, and charts. Applies per-sheet normalization and global rescaling.
    /// </summary>
    /// <param name="wbKey">The key workbook.</param>
    /// <param name="sheetHint">Optional sheet name or substring hint.</param>
    /// <param name="allSheets">If true, process all sheets; otherwise use the hint/fallbacks.</param>
    /// <param name="targetTotal">Desired total points to scale to (≤0 to skip scaling).</param>
    /// <param name="keyZipBytes">Optional XLSX ZIP bytes for artifact scanners; if null, scanners are skipped.</param>
    /// <returns>The generated <see cref="Rubric"/>.</returns>
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

    /// <summary>
    /// Appends pivot layout rules extracted from workbook ZIP to an existing check list,
    /// avoiding duplicates by pivot table name (case-insensitive).
    /// </summary>
    /// <param name="checks">Destination rule list for the sheet.</param>
    /// <param name="zipRules">Pivot rules obtained from ZIP parsing.</param>
    private static void MergePivotRules(List<Rule> checks, List<Rule> zipRules)
    {
        var existing = new HashSet<string>(
            checks.Where(c => c.Type == "pivot_layout" && c.Pivot?.TableNameLike != null)
                  .Select(c => c.Pivot!.TableNameLike!), StringComparer.OrdinalIgnoreCase);

        foreach (var r in zipRules)
            if (r.Pivot?.TableNameLike is null || !existing.Contains(r.Pivot.TableNameLike))
                checks.Add(r);
    }

    /// <summary>
    /// Builds a rubric from explicit per-sheet sections and ranges.
    /// Generates per-cell formula/value checks, summary formulas, intersecting table rules,
    /// and (optionally) artifacts discovered from the ZIP. Preserves the provided section order.
    /// </summary>
    /// <param name="wbKey">The key workbook.</param>
    /// <param name="sectionsPerSheet">Map of sheet name → list of (section name, A1 ranges).</param>
    /// <param name="includeArtifacts">If true, append discovered pivot/CF/chart rules under an “Artifacts” section.</param>
    /// <param name="targetTotal">Desired total points (≤0 to skip scaling).</param>
    /// <param name="keyZipBytes">Optional XLSX ZIP bytes for artifact discovery.</param>
    /// <returns>The generated <see cref="Rubric"/>.</returns>
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

                // (4) Excel Tables that intersect the selected ranges
                foreach (var tbl in ws.Tables)
                {
                    var tAddr = tbl.RangeAddress;

                    // Any intersection with the user-selected ranges?
                    bool intersects = addrs.Any(ra =>
                    {
                        var a1 = ra.FirstAddress; var a2 = ra.LastAddress;
                        var b1 = tAddr.FirstAddress; var b2 = tAddr.LastAddress;
                        return !(a2.RowNumber < b1.RowNumber ||
                                 b2.RowNumber < a1.RowNumber ||
                                 a2.ColumnNumber < b1.ColumnNumber ||
                                 b2.ColumnNumber < a1.ColumnNumber);
                    });
                    if (!intersects) continue;

                    // Avoid duplicates if another section already added the same table
                    bool already = checks.Any(c => c.Type == "table" &&
                                                   string.Equals(c.Table?.NameLike, tbl.Name, StringComparison.OrdinalIgnoreCase));
                    if (already) continue;

                    // Columns (headers)
                    var cols = tbl.Fields
                        .Select(f => (f.Name ?? string.Empty).Trim())
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .ToList();
                    if (cols.Count == 0) continue;

                    var fullRangeA1 = tAddr.ToStringRelative();
                    var body = tbl.DataRange;
                    int bodyRows = body?.RowCount() ?? 0;
                    int bodyCols = body?.ColumnCount() ?? 0;

                    var bodyVals = new List<List<string>>();
                    if (body != null)
                    {
                        foreach (var r in body.Rows())
                        {
                            var rowVals = new List<string>();
                            foreach (var c in r.Cells())
                                rowVals.Add(c.GetFormattedString() ?? string.Empty);
                            bodyVals.Add(rowVals);
                        }
                    }

                    checks.Add(new Rule
                    {
                        Type = "table",
                        Points = 1.0,
                        Section = sectionName,
                        Note = $"Table '{tbl.Name}' columns",
                        Table = new TableSpec
                        {
                            Sheet = ws.Name,
                            NameLike = tbl.Name,

                            // headers
                            Columns = cols,
                            RequireOrder = false,

                            // location & size
                            RangeRef = fullRangeA1,
                            Rows = bodyRows,
                            Cols = bodyCols,

                            // defaults (UI can toggle these)
                            AllowExtraRows = null,
                            AllowExtraCols = null,

                            // captured data so the UI can show it
                            BodyMatch = null,
                            BodyOrderMatters = null,
                            BodyCaseSensitive = null,
                            BodyTrim = true,
                            BodyRows = bodyVals,

                            // empty by default; UI can add “must contain” rows later
                            ContainsRows = new List<Dictionary<string, string>>()
                        }
                    });
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

    /// <summary>
    /// Builds a rubric from a map of sheet → ranges (no explicit sections).
    /// Generates per-cell rules, plus optional artifacts extracted from the provided ZIP.
    /// </summary>
    /// <param name="wbKey">The key workbook.</param>
    /// <param name="ranges">Map of sheet name → A1 ranges to harvest.</param>
    /// <param name="includeArtifacts">If true, include pivot/CF/chart rules discovered via ZIP.</param>
    /// <param name="targetTotal">Desired total points (≤0 to skip scaling).</param>
    /// <param name="keyZipBytes">Optional XLSX ZIP bytes for artifact discovery.</param>
    /// <returns>The generated <see cref="Rubric"/>.</returns>
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
