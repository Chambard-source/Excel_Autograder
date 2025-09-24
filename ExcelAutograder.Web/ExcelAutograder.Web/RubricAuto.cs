using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;
using System.Text.Json;

public static class RubricAuto
{
    /// Back-compat: build and scale to a sensible default (5.0).
    public static Rubric GenerateFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets)
    {
        const double defaultTotal = 5.0;
        // No raw bytes here (legacy path) → ZIP pivot fallback will be skipped.
        return BuildFromKey(wbKey, sheetHint, allSheets, defaultTotal, keyZipBytes: null);
    }

    /// Scale all rule points so the rubric totals to the requested target.
    public static void ScalePoints(Rubric rub, double target) => RescalePoints(rub, target);

    /// Build from a key workbook (optionally 1 sheet by hint or all sheets) and then scale to targetTotal.
    /// Legacy overload (no raw bytes) — still works; just delegates to new overload.
    public static Rubric BuildFromKey(XLWorkbook wbKey, string? sheetHint, bool allSheets, double targetTotal)
        => BuildFromKey(wbKey, sheetHint, allSheets, targetTotal, keyZipBytes: null);

    /// New overload: same as above, but accepts the raw .xlsx bytes so we can ZIP-scan pivots
    /// when ClosedXML doesn’t expose them.
    public static Rubric BuildFromKey(
        XLWorkbook wbKey,
        string? sheetHint,
        bool allSheets,
        double targetTotal,
        byte[]? keyZipBytes)
    {
        // Pre-scan pivots from the ZIP for all sheets (works when ClosedXML doesn't expose them)
        var zipPivots = keyZipBytes is null
            ? new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase)
            : ExtractPivotRulesFromZip(keyZipBytes);

        var zipCF = keyZipBytes is null
            ? new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase)
            : ExtractConditionalRulesFromZip(keyZipBytes);

        var zipCharts = keyZipBytes is null
            ? new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase)
            : ExtractChartRulesFromZip(keyZipBytes);

        var rub = new Rubric
        {
            Points = 0,
            Sheets = new Dictionary<string, SheetSpec>(StringComparer.OrdinalIgnoreCase),
            Report = new Report { IncludePassFailColumn = true, IncludeComments = true }
        };

        var sheets = ResolveSheets(wbKey, sheetHint, allSheets).ToList();
        if (sheets.Count == 0 && wbKey.Worksheets.Count > 0)
            sheets.Add(wbKey.Worksheets.First());

        foreach (var ws in sheets)
        {
            var checks = BuildChecksForSheet(ws);

            // Merge in ZIP-discovered pivot rules for this sheet (if any)
            if (zipPivots.TryGetValue(ws.Name, out var prules) && prules.Count > 0)
            {
                var existingNames = new HashSet<string>(
                    checks.Where(c => c.Type == "pivot_layout" && c.Pivot?.TableNameLike != null)
                          .Select(c => c.Pivot!.TableNameLike!),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var r in prules)
                {
                    var name = r.Pivot?.TableNameLike;
                    if (name == null || !existingNames.Contains(name))
                        checks.Add(r);
                }
            }

            if (zipCF.TryGetValue(ws.Name, out var cfRules) && cfRules.Count > 0)
                checks.AddRange(cfRules);

            if (zipCharts.TryGetValue(ws.Name, out var crules) && crules.Count > 0)
                checks.AddRange(crules);

            if (checks.Count == 0)
                checks.Add(new Rule { Type = "custom_note", Points = 0, Note = "No auto-generated checks for this sheet" });

            rub.Sheets[ws.Name] = new SheetSpec { Checks = checks };
            rub.Points += checks.Sum(c => c.Points);
        }

        foreach (var kv in zipCharts)
        {
            if (!rub.Sheets.ContainsKey(kv.Key))
            {
                rub.Sheets[kv.Key] = new SheetSpec { Checks = kv.Value };
                rub.Points += kv.Value.Sum(r => r.Points);
            }
        }

        if (targetTotal > 0) RescalePoints(rub, targetTotal);
        return rub;
    }

    // ---------------------------
    // Internals / helpers
    // ---------------------------

    private static IEnumerable<IXLWorksheet> ResolveSheets(XLWorkbook wb, string? hint, bool all)
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

    /// Heuristics + pivot discovery per sheet.
    private static List<Rule> BuildChecksForSheet(IXLWorksheet ws)
    {
        var checks = new List<Rule>();
        var used = ws.RangeUsed();
        if (used == null)
        {
            // still try to append pivot rules (some sheets may have only a pivot)
            checks.AddRange(BuildPivotRulesForSheet(ws));
            if (checks.Count == 0)
                checks.Add(new Rule { Type = "custom_note", Points = 0, Note = "No auto-generated checks for this sheet" });
            return checks;
        }

        // 1) Header formats A1 / B1 if present
        TryAddHeaderFormat(ws, "A1", "Header Student bold / size", checks);
        TryAddHeaderFormat(ws, "B1", "Header Score bold / size", checks);

        // 2) Simple data block guess (cols A/B from row 2)
        var maxRow = Math.Max(used.RangeAddress.LastAddress.RowNumber, 12);
        int lastDataRowA = FindLastDataRow(ws, "A", 2, maxRow);
        int lastDataRowB = FindLastDataRow(ws, "B", 2, maxRow);

        bool colAIsPrimary = lastDataRowA >= lastDataRowB;
        string primCol = colAIsPrimary ? "A" : "B";
        string secCol = colAIsPrimary ? "B" : "A";
        int lastRow = Math.Max(lastDataRowA, lastDataRowB);
        if (lastRow < 2) lastRow = 2;

        // 3) Sequence check in primary column (1..n)
        if (lastRow >= 2 && LooksLikeSequence(ws, primCol, 2, lastRow))
        {
            checks.Add(new Rule
            {
                Type = "range_sequence",
                Points = 0.5,
                Range = $"{primCol}2:{primCol}{lastRow}",
                Note = $"Column {primCol} is 1..{lastRow - 1}",
                Start = 1,
                Step = 1
            });
        }

        // 4) Numeric column
        if (lastRow >= 2 && LooksNumeric(ws, secCol, 2, lastRow))
        {
            checks.Add(new Rule
            {
                Type = "range_numeric",
                Points = 0.5,
                Range = $"{secCol}2:{secCol}{lastRow}",
                Note = $"Column {secCol} contains numbers"
            });

            // 5) Number format (most frequent)
            var fmt = GetRangeNumberFormat(ws, $"{secCol}2:{secCol}{lastRow}");
            if (!string.IsNullOrWhiteSpace(fmt))
            {
                checks.Add(new Rule
                {
                    Type = "range_format",
                    Points = 0.25,
                    Range = $"{secCol}2:{secCol}{lastRow}",
                    Note = $"Column {secCol} number format",
                    Format = new FormatSpec { NumberFormat = fmt }
                });
            }
        }

        // 6) Bottom summary in numeric column (row after data)
        int summaryRow = lastRow + 1;
        var summaryAddr = $"{secCol}{summaryRow}";
        var summaryCell = ws.Cell(summaryAddr);
        var sf = NormalizeFormula(summaryCell.FormulaA1);

        if (!string.IsNullOrWhiteSpace(sf) && ContainsAny(sf, "AVERAGE", "SUM", "COUNT", "COUNTA", "MIN", "MAX"))
        {
            checks.Add(new Rule
            {
                Type = "formula",
                Points = 1.5,
                Cell = summaryAddr,
                Note = "Summary formula",
                ExpectedFromKey = true,
                Expected = JsonSerializer.SerializeToElement(KeyCellExpectedText(summaryCell))
            });

            // If the key’s summary uses absolutes, require a $ reference to be present (regex)
            if (HasAbsoluteRef(sf))
            {
                checks.Add(new Rule
                {
                    Type = "formula",
                    Points = 0.5,
                    Cell = summaryAddr,
                    Note = "Summary uses absolute reference(s)",
                    AllowRegex = true,
                    ExpectedFormulaRegex = @".*\$[A-Za-z]+\$?\d.*"
                });
            }
        }

        // 7) Border outline A{n+1}:B{n+1}
        var borderRange = $"A{summaryRow}:B{summaryRow}";
        if (HasOutlineBorder(ws, borderRange))
        {
            checks.Add(new Rule
            {
                Type = "range_format",
                Points = 1.0,
                Range = borderRange,
                Note = $"Borders around {borderRange}",
                Format = new FormatSpec { Border = new BorderSpec { Outline = true } }
            });
        }

        // 7.5) Tables on this sheet → table rules
        foreach (var tbl in ws.Tables)
        {
            // header names
            var cols = tbl.Fields
                .Select(f => (f.Name ?? string.Empty).Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
            if (cols.Count == 0) continue;

            // A1 style range that includes header + data area
            var fullRangeA1 = tbl.RangeAddress.ToStringRelative();

            // body (data only) dimensions
            var body = tbl.DataRange;
            int bodyRows = body?.RowCount() ?? 0;
            int bodyCols = body?.ColumnCount() ?? 0;

            // capture the formatted body values (so we can optionally grade contents)
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
                Note = $"Table '{tbl.Name}' columns",
                Table = new TableSpec
                {
                    Sheet = ws.Name,
                    NameLike = tbl.Name,

                    // column headers (order match optional)
                    Columns = cols,
                    RequireOrder = false,

                    // range + dimensions
                    RangeRef = fullRangeA1,
                    Rows = bodyRows,
                    Cols = bodyCols,

                    // dimension flexibility defaults (UI can toggle to true)
                    AllowExtraRows = null,
                    AllowExtraCols = null,

                    // content checks (all optional; UI can toggle BodyMatch/…)
                    BodyMatch = null,          // if true, check contents against BodyRows
                    BodyOrderMatters = null,   // if true, row order must match
                    BodyCaseSensitive = null,  // if true, compare strings case-sensitively
                    BodyTrim = true,           // trim cell text before compare
                    BodyRows = bodyVals,       // the captured body we’ll grade against

                    // “must contain at least these rows” (starts empty; UI can add)
                    ContainsRows = new List<Dictionary<string, string>>()
                }
            });
        }

        // 8) Pick up *all other* formulas in the used range
        var alreadyCovered = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            summaryAddr // avoid duplicating the summary cell
        };

        foreach (var c in used.CellsUsed(XLCellsUsedOptions.All))
        {
            var a1 = c.Address.ToString();
            var f = NormalizeFormula(c.FormulaA1);
            if (string.IsNullOrWhiteSpace(f)) continue;

            var r = new Rule
            {
                Type = "formula",
                Points = 0.5,                       // UI will rescale later
                Cell = a1,
                Note = "Formula from key",
                ExpectedFromKey = true,
                ExpectedFormula = f,                // UI can display it for review
                Expected = JsonSerializer.SerializeToElement(KeyCellExpectedText(c))
            };
            if (HasAbsoluteRef(f))
                r.RequireAbsolute = true;

            checks.Add(r);
        }

        // --- Pivot rules ---
        // First try ClosedXML reflection (works on many files)
        checks.AddRange(BuildPivotRulesForSheet(ws));

        // If none found, fall back to presence heuristics (useful for static “pivot-like” builds)
        if (!checks.Any(r => r.Type == "pivot_layout"))
            checks.AddRange(BuildPivotPresenceHeuristics(ws));

        if (checks.Count == 0)
            checks.Add(new Rule { Type = "custom_note", Points = 0, Note = "No auto-generated checks for this sheet" });

        return checks;
    }

    private static List<Rule> BuildPivotRulesForSheet(IXLWorksheet ws)
    {
        var rules = new List<Rule>();

        // Access ws.PivotTables via reflection (version-agnostic)
        var pivotsObj = ws.GetType().GetProperty("PivotTables")?.GetValue(ws);
        var pivots = AsEnum(pivotsObj);
        if (!pivots.Any()) return rules;

        foreach (var pt in pivots)
        {
            var ptName = GetStrProp(pt, "Name") ?? ws.Name + " Pivot";

            // Collect layout parts
            var rows = GetEnumProp(pt, "RowLabels")
                .Select(f => FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")))
                .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var cols = GetEnumProp(pt, "ColumnLabels")
                .Select(f => FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")))
                .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var filters = GetEnumProp(pt, "ReportFilters")
                .Select(f => FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")))
                .Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var values = new List<PivotValueSpec>();
            foreach (var v in GetEnumProp(pt, "Values"))
            {
                var field = FirstNonEmpty(GetStrProp(v, "SourceName"), GetStrProp(v, "CustomName"), GetStrProp(v, "Name"));
                if (string.IsNullOrWhiteSpace(field)) continue;

                var sf = GetStrProp(v, "SummaryFormula") ?? GetStrProp(v, "Function") ?? "";
                values.Add(new PivotValueSpec { Field = field, Agg = NormAgg(sf) });
            }

            // Skip if layout is empty (defensive)
            if (rows.Count == 0 && cols.Count == 0 && filters.Count == 0 && values.Count == 0)
                continue;

            // Create a pivot_layout rule (points are provisional; will be scaled later)
            rules.Add(new Rule
            {
                Type = "pivot_layout",
                Points = 1.5,
                Note = $"Pivot '{ptName}' layout",
                Pivot = new PivotSpec
                {
                    Sheet = ws.Name,
                    TableNameLike = ptName,
                    Rows = rows.Count > 0 ? rows : null,
                    Columns = cols.Count > 0 ? cols : null,
                    Filters = filters.Count > 0 ? filters : null,
                    Values = values.Count > 0 ? values : null
                }
            });
        }

        return rules;
    }

    private static IEnumerable<Rule> BuildPivotPresenceHeuristics(IXLWorksheet ws)
    {
        // Quick presence rules when there's no real pivot object
        var rules = new List<Rule>();
        var hits = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var t in ws.Tables)
            if (!string.IsNullOrWhiteSpace(t.Name) && t.Name.Contains("Pivot", StringComparison.OrdinalIgnoreCase))
                hits.Add(t.Name);

        foreach (var nr in ws.NamedRanges)
            if (!string.IsNullOrWhiteSpace(nr.Name) && nr.Name.Contains("Pivot", StringComparison.OrdinalIgnoreCase))
                hits.Add(nr.Name);

        foreach (var c in ws.CellsUsed(XLCellsUsedOptions.All))
        {
            var s = c.GetString();
            if (!string.IsNullOrWhiteSpace(s) && s.Contains("Pivot", StringComparison.OrdinalIgnoreCase))
                hits.Add(s);
        }

        foreach (var name in hits)
        {
            rules.Add(new Rule
            {
                Type = "custom_note",
                Points = 0.5,
                Note = $"Pivot-like '{name}' exists",
                Require = new RequireSpec { Sheet = ws.Name, PivotTableLike = name }
            });
        }

        return rules;
    }

    // -------- ZIP (.xlsx) pivot discovery --------
    private static Dictionary<string, List<Rule>> ExtractPivotRulesFromZip(byte[] keyZipBytes)
    {
        // sheetName -> rules
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(keyZipBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);

        // Map sheet index to sheet name from xl/workbook.xml
        var wbEntry = zip.GetEntry("xl/workbook.xml");
        if (wbEntry is null) return map;

        var wbXml = XDocument.Load(wbEntry.Open());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        var sheetIndexToName = new Dictionary<int, string>();
        int idx = 1;
        var sheetsEl = wbXml.Root?.Element(ns + "sheets");
        if (sheetsEl != null)
        {
            foreach (var sh in sheetsEl.Elements(ns + "sheet"))
            {
                var n = (string?)sh.Attribute("name") ?? $"Sheet{idx}";
                sheetIndexToName[idx] = n;
                idx++;
            }
        }

        // For each sheetN.xml, open its _rels to find pivotTable targets
        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsEntry = zip.GetEntry(relsPath);
            if (relsEntry is null) continue;

            var relsXml = XDocument.Load(relsEntry.Open());
            var rels = relsXml.Root!.Elements()
                        .Where(e => ((string?)e.Attribute("Type"))?.Contains("pivotTable", StringComparison.OrdinalIgnoreCase) == true)
                        .Select(e => (string?)e.Attribute("Target"))
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();


            foreach (var target in rels)
            {
                var normalized = target!.Replace("\\", "/");
                var full = target.StartsWith("/") ? target.TrimStart('/')
                         : target.StartsWith("../") ? "xl/" + target.Replace("../", "")
                         : "xl/worksheets/" + target;

                if (full.Contains("/pivotTables/", StringComparison.OrdinalIgnoreCase))
                {
                    var pos = full.IndexOf("/pivotTables/", StringComparison.OrdinalIgnoreCase);
                    full = "xl" + full.Substring(pos);
                }

                var ptEntry = zip.GetEntry(full);
                if (ptEntry is null) continue;

                var ptXml = XDocument.Load(ptEntry.Open());
                var def = ptXml.Root!;
                var ptName = (string?)def.Attribute("name") ?? "Pivot";

                // Gather pivotFields to map indexes → names
                var pivotFields = def.Element(ns + "pivotFields")?.Elements(ns + "pivotField").ToList()
                                   ?? new List<XElement>();

                string NameByIndex(int fi)
                {
                    if (fi >= 0 && fi < pivotFields.Count)
                    {
                        var pf = pivotFields[fi];
                        var nAttr = (string?)pf.Attribute("name");
                        if (!string.IsNullOrWhiteSpace(nAttr)) return nAttr!;
                    }
                    return $"Field{fi}";
                }

                var rows = new List<string>();
                foreach (var rf in def.Element(ns + "rowFields")?.Elements(ns + "field") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)rf.Attribute("x"), out var fi)) rows.Add(NameByIndex(fi));

                var cols = new List<string>();
                foreach (var cf in def.Element(ns + "colFields")?.Elements(ns + "field") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)cf.Attribute("x"), out var fi)) cols.Add(NameByIndex(fi));

                var filters = new List<string>();
                foreach (var pf in def.Element(ns + "pageFields")?.Elements(ns + "pageField") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)pf.Attribute("fld"), out var fi)) filters.Add(NameByIndex(fi));

                var values = new List<PivotValueSpec>();
                foreach (var df in def.Element(ns + "dataFields")?.Elements(ns + "dataField") ?? Enumerable.Empty<XElement>())
                {
                    var nm = (string?)df.Attribute("name");
                    var fld = (string?)df.Attribute("fld");
                    string fieldName = !string.IsNullOrWhiteSpace(nm) ? nm!
                                      : (int.TryParse(fld, out var fi) ? NameByIndex(fi) : "Value");

                    var subtotal = ((string?)df.Attribute("subtotal"))?.ToLowerInvariant() ?? "sum";
                    string agg = subtotal.Contains("count") ? "count"
                               : subtotal.Contains("average") || subtotal.Contains("avg") ? "average"
                               : subtotal.Contains("min") ? "min"
                               : subtotal.Contains("max") ? "max"
                               : "sum";

                    values.Add(new PivotValueSpec { Field = fieldName, Agg = agg });
                }

                var rule = new Rule
                {
                    Type = "pivot_layout",
                    Points = 1.5,
                    Note = $"Pivot '{ptName}' layout",
                    Pivot = new PivotSpec
                    {
                        Sheet = sheetName,
                        TableNameLike = ptName,
                        Rows = rows.Count > 0 ? rows.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Columns = cols.Count > 0 ? cols.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Filters = filters.Count > 0 ? filters.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Values = values.Count > 0 ? values : null
                    }
                };

                if (!map.TryGetValue(sheetName, out var list)) map[sheetName] = list = new List<Rule>();
                list.Add(rule);
            }
        }

        return map;
    }

    private static Dictionary<string, List<Rule>> ExtractConditionalRulesFromZip(byte[] keyZipBytes)
    {
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(keyZipBytes);
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        var wbEntry = zip.GetEntry("xl/workbook.xml"); if (wbEntry is null) return map;
        var wbXml = System.Xml.Linq.XDocument.Load(wbEntry.Open());
        XName S(string n) => System.Xml.Linq.XName.Get(n, "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        var idxToName = new Dictionary<int, string>(); int idx = 1;
        var sheetsEl = wbXml.Root?.Element(S("sheets"));
        if (sheetsEl != null)
            foreach (var sh in sheetsEl.Elements(S("sheet")))
                idxToName[idx++] = (string?)sh.Attribute("name") ?? $"Sheet{idx}";

        // styles for dxf (optional)
        var stylesEntry = zip.GetEntry("xl/styles.xml");
        System.Xml.Linq.XDocument? stylesXml = stylesEntry != null ? System.Xml.Linq.XDocument.Load(stylesEntry.Open()) : null;

        for (int i = 1; i < idx; i++)
        {
            if (!idxToName.TryGetValue(i, out var sheetName)) continue;
            var sheetPath = $"xl/worksheets/sheet{i}.xml";
            var sheetEntry = zip.GetEntry(sheetPath); if (sheetEntry is null) continue;
            var wsXml = System.Xml.Linq.XDocument.Load(sheetEntry.Open());

            foreach (var block in wsXml.Root!.Elements(S("conditionalFormatting")))
            {
                var sqref = (string?)block.Attribute("sqref") ?? "";
                foreach (var ruleEl in block.Elements(S("cfRule")))
                {
                    var t = (string?)ruleEl.Attribute("type");
                    var op = (string?)ruleEl.Attribute("operator");
                    var frms = ruleEl.Elements(S("formula")).Select(e => e.Value?.Trim()).ToList();
                    var txt = (string?)ruleEl.Attribute("text");

                    string? fillRgb = null;
                    var dxfIdAttr = ruleEl.Attribute("dxfId");
                    if (stylesXml != null && dxfIdAttr != null && int.TryParse(dxfIdAttr.Value, out var dxfId))
                    {
                        var dxfs = stylesXml.Root?.Element(S("dxfs"))?.Elements(S("dxf")).ToList();
                        if (dxfs != null && dxfId >= 0 && dxfId < dxfs.Count)
                        {
                            var dxf = dxfs[dxfId];
                            fillRgb = dxf.Element(S("fill"))?.Element(S("patternFill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value
                                   ?? dxf.Element(S("fill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value;
                            if (!string.IsNullOrWhiteSpace(fillRgb) && fillRgb.Length == 8) // often ARGB like FFxxxxxx
                                fillRgb = fillRgb.Substring(2);
                        }
                    }

                    var rule = new Rule
                    {
                        Type = "conditional_format",
                        Points = 0.5, // will be rescaled later
                        Note = "Conditional format",
                        Cond = new ConditionalFormatSpec
                        {
                            Sheet = sheetName,
                            Range = sqref.Split(' ').FirstOrDefault(), // first target range
                            Type = t,
                            Operator = MapXmlOp(op),
                            Formula1 = frms.ElementAtOrDefault(0),
                            Formula2 = frms.ElementAtOrDefault(1),
                            Text = txt,
                            FillRgb = fillRgb
                        }
                    };

                    if (!map.TryGetValue(sheetName, out var list)) map[sheetName] = list = new List<Rule>();
                    list.Add(rule);
                }
            }
        }
        return map;
    }

    static string? MapXmlOp(string? op) => op switch
    {
        "greaterThan" => "gt",
        "greaterThanOrEqual" => "ge",
        "lessThan" => "lt",
        "lessThanOrEqual" => "le",
        "equal" => "eq",
        "notEqual" => "ne",
        "between" => "between",
        "notBetween" => "notBetween",
        _ => op
    };


    private static void TryAddHeaderFormat(IXLWorksheet ws, string cellAddr, string note, List<Rule> checks)
    {
        var c = ws.Cell(cellAddr);
        if (c == null) return;

        bool isBold = c.Style.Font.Bold;
        double size = c.Style.Font.FontSize;

        if (isBold || size > 11.0)
        {
            checks.Add(new Rule
            {
                Type = "format",
                Points = 0.125,
                Cell = cellAddr,
                Note = note,
                Format = new FormatSpec
                {
                    Bold = isBold ? true : null,
                    Font = new FontSpec { Size = size }
                }
            });
        }
    }

    private static int FindLastDataRow(IXLWorksheet ws, string col, int startRow, int maxRow)
    {
        int last = startRow - 1;
        for (int r = startRow; r <= maxRow; r++)
        {
            var cell = ws.Cell($"{col}{r}");
            var v = cell.Value;

            if (cell.IsEmpty() || string.IsNullOrWhiteSpace(v.ToString()))
                break;

            last = r;
        }
        return last;
    }

    private static bool LooksLikeSequence(IXLWorksheet ws, string col, int r1, int r2)
    {
        int expected = 1;
        bool any = false;
        for (int r = r1; r <= r2; r++)
        {
            var v = ws.Cell($"{col}{r}").Value;
            if (!TryToInt(v, out var n)) return false;
            if (n != expected) return false;
            expected++;
            any = true;
        }
        return any;
    }

    private static bool LooksNumeric(IXLWorksheet ws, string col, int r1, int r2)
    {
        int count = 0, ok = 0;
        for (int r = r1; r <= r2; r++)
        {
            count++;
            var v = ws.Cell($"{col}{r}").Value;
            if (TryToDouble(v, out _)) ok++;
        }
        return count > 0 && ok >= Math.Max(1, (int)Math.Ceiling(count * 0.7));
    }

    private static string? GetRangeNumberFormat(IXLWorksheet ws, string rangeA1)
    {
        var rng = ws.Range(rangeA1);
        var fmts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var c in rng.CellsUsed())
        {
            var f = c.Style.NumberFormat.Format ?? string.Empty;
            if (string.IsNullOrWhiteSpace(f)) continue;
            fmts.TryGetValue(f, out var n);
            fmts[f] = n + 1;
        }
        if (fmts.Count == 0) return null;
        return fmts.OrderByDescending(kv => kv.Value).First().Key;
    }

    private static bool HasOutlineBorder(IXLWorksheet ws, string rangeA1)
    {
        var rng = ws.Range(rangeA1);
        var b = rng.Style.Border;
        return b.LeftBorder != XLBorderStyleValues.None
            || b.RightBorder != XLBorderStyleValues.None
            || b.TopBorder != XLBorderStyleValues.None
            || b.BottomBorder != XLBorderStyleValues.None;
    }

    private static bool ContainsAny(string text, params string[] needles)
        => needles.Any(n => text.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0);

    private static bool HasAbsoluteRef(string formula) => formula.IndexOf('$') >= 0;

    private static string NormalizeFormula(string? f)
    {
        var s = f?.Trim() ?? string.Empty;
        if (s.Length > 0 && s[0] != '=') s = "=" + s;
        return s;
    }

    private static string? KeyCellExpectedText(IXLCell kc)
    {
        // 1) Prefer what Excel displays (honors number formats, % etc.)
        var s = kc.GetFormattedString();
        if (!string.IsNullOrWhiteSpace(s)) return s;

        // 2) Fall back to an invariant representation by data type
        switch (kc.DataType)
        {
            case XLDataType.Number:
                return kc.GetValue<double>().ToString(CultureInfo.InvariantCulture);

            case XLDataType.Boolean:
                return kc.GetValue<bool>() ? "TRUE" : "FALSE";

            case XLDataType.DateTime:
                return kc.GetValue<DateTime>().ToString(CultureInfo.InvariantCulture);

            case XLDataType.Text:
                return kc.GetString();

            case XLDataType.TimeSpan:
                return kc.GetValue<TimeSpan>().ToString();

            case XLDataType.Blank:
                return null; // nothing to show

            default:
                // XLCellValue is a struct (never null) → no ?.
                return kc.Value.ToString();
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

    // --- Reflection helpers (ClosedXML version–agnostic)
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
    public static void RescalePoints(Rubric rub, double desiredTotal)
    {
        if (rub == null) return;

        double sum = 0;
        if (rub.Sheets != null)
        {
            foreach (var sheet in rub.Sheets.Values)
            {
                if (sheet?.Checks == null) continue;
                foreach (var rule in sheet.Checks) sum += rule.Points;
            }
        }

        if (sum <= 0) { rub.Points = 0; return; }

        double k = desiredTotal / sum;
        foreach (var sheet in rub.Sheets.Values)
        {
            if (sheet?.Checks == null) continue;
            foreach (var rule in sheet.Checks)
                rule.Points = Math.Round(rule.Points * k, 3);
        }
        rub.Points = desiredTotal;
    }

    // --------------------- CHART RULES (auto from key) ---------------------
    private class AutoChartInfo
    {
        public string Sheet = "";
        public string Name = "Chart";
        public string Type = ""; // line/column/bar/pie/scatter/area/doughnut/radar/bubble
        public string? Title, TitleRef;
        public string? LegendPos;
        public bool DataLabels;
        public string? XTitle, YTitle;
        public List<AutoSeriesInfo> Series = new();
    }
    private class AutoSeriesInfo
    {
        public string? Name, NameRef, CatRef, ValRef;
    }

    private static Dictionary<string, List<Rule>> ExtractChartRulesFromZip(byte[] keyZipBytes)
    {
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);
        var chartsBySheet = ParseChartsFromZipAuto(keyZipBytes);

        foreach (var kv in chartsBySheet)
        {
            var sheet = kv.Key;
            foreach (var ch in kv.Value)
            {
                var rule = new Rule
                {
                    Type = "chart",
                    Points = 1.5, // UI rescaled later
                    Note = $"Chart '{ch.Name}' on {sheet}",
                    Chart = new ChartSpec
                    {
                        Sheet = sheet,
                        NameLike = ch.Name,
                        Type = ch.Type,
                        Title = ch.Title,
                        TitleRef = ch.TitleRef,
                        LegendPos = ch.LegendPos,
                        DataLabels = ch.DataLabels,
                        XTitle = ch.XTitle,
                        YTitle = ch.YTitle,
                        Series = ch.Series.Select(s => new ChartSeriesSpec
                        {
                            Name = s.Name,
                            NameRef = s.NameRef,
                            CatRef = s.CatRef,
                            ValRef = s.ValRef
                        }).ToList()
                    }
                };

                if (!map.TryGetValue(sheet, out var list)) map[sheet] = list = new();
                list.Add(rule);
            }
        }

        return map;
    }

    // ---- minimal OOXML parser (local; no ClosedXML dependency) ----
    private static Dictionary<string, List<AutoChartInfo>> ParseChartsFromZipAuto(byte[] zipBytes)
    {
        var result = new Dictionary<string, List<AutoChartInfo>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(zipBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);

        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace pkg = "http://schemas.openxmlformats.org/package/2006/relationships";

        string ReadEntryText(string path)
        {
            var e = zip.GetEntry(path);
            if (e == null) return "";
            using var s = e.Open();
            using var r = new StreamReader(s);
            return r.ReadToEnd();
        }

        // 1) map sheet index -> sheet name (sheet1.xml → "Summary", etc.)
        var wbTxt = ReadEntryText("xl/workbook.xml");
        if (string.IsNullOrEmpty(wbTxt)) return result;
        var wb = XDocument.Parse(wbTxt);

        var sheetIndexToName = new Dictionary<int, string>();
        var sheetsEl = wb.Root?.Element(nsMain + "sheets");
        int idx = 1;
        if (sheetsEl != null)
        {
            foreach (var sh in sheetsEl.Elements(nsMain + "sheet"))
            {
                var nm = (string?)sh.Attribute("name") ?? $"Sheet{idx}";
                sheetIndexToName[idx++] = nm;
            }
        }

        // 2) for each sheet rels → drawing → chart
        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsTxt = ReadEntryText(relsPath);
            if (string.IsNullOrEmpty(relsTxt)) continue;

            var rels = XDocument.Parse(relsTxt);
            var drawingTargets = rels.Root?
                .Elements(pkg + "Relationship")  // <— was rel + "Relationship"
                .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/drawing") == true)
                .Select(r => ((string?)r.Attribute("Target"))?.TrimStart('/').Replace("../", "xl/"))
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList() ?? new List<string>();

            foreach (var drawingTarget in drawingTargets)
            {
                var drPath = drawingTarget!.StartsWith("xl/") ? drawingTarget : $"xl/{drawingTarget}";
                var drTxt = ReadEntryText(drPath);
                if (string.IsNullOrEmpty(drTxt)) continue;
                var drXml = XDocument.Parse(drTxt);

                // drawing rels
                var drRelsPath = drPath.Replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
                var drRelsTxt = ReadEntryText(drRelsPath);
                if (string.IsNullOrEmpty(drRelsTxt)) continue;
                var drRels = XDocument.Parse(drRelsTxt);

                // find c:chart r:id within graphicFrames
                var frames = drXml.Descendants(xdr + "graphicFrame").ToList();
                foreach (var gf in frames)
                {
                    var cNvPr = gf.Element(xdr + "nvGraphicFramePr")?.Element(xdr + "cNvPr");
                    var frameName = cNvPr?.Attribute("name")?.Value ?? "Chart";

                    var chartElem = gf.Descendants(a + "graphicData").Descendants(c + "chart").FirstOrDefault();
                    var rid = chartElem?.Attribute(rel + "id")?.Value;
                    if (string.IsNullOrWhiteSpace(rid)) continue;

                    var target = drRels.Root?
                        .Elements(pkg + "Relationship")  // <— was rel + "Relationship"
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;

                    if (string.IsNullOrWhiteSpace(target)) continue;

                    var tgt = (target ?? "").Replace("\\", "/");
                    var chartPath = tgt.StartsWith("/") ? tgt.TrimStart('/')
                                 : tgt.StartsWith("../") ? "xl/" + tgt.Substring(3)
                                 : tgt.StartsWith("xl/") ? tgt : "xl/" + tgt;

                    var chTxt = ReadEntryText(chartPath);
                    if (string.IsNullOrEmpty(chTxt)) continue;

                    var chDoc = XDocument.Parse(chTxt);
                    var info = new AutoChartInfo { Sheet = sheetName, Name = frameName };

                    var plotArea = chDoc.Descendants(c + "plotArea").FirstOrDefault();
                    info.Type = DetectChartType(plotArea, c);   // or DetectChartType(plotArea) if your helper takes 1 arg

                    // Title
                    var titleEl = chDoc.Descendants(c + "title").FirstOrDefault();
                    (info.Title, info.TitleRef) = ReadChartTextAuto(titleEl, c, a);

                    // Axis titles
                    var catAx = plotArea?.Elements(c + "catAx").FirstOrDefault();
                    var valAx = plotArea?.Elements(c + "valAx").FirstOrDefault();
                    (info.XTitle, _) = ReadChartTextAuto(catAx?.Element(c + "title"), c, a);
                    (info.YTitle, _) = ReadChartTextAuto(valAx?.Element(c + "title"), c, a);

                    // Legend + labels
                    var leg = chDoc.Descendants(c + "legend").FirstOrDefault();
                    info.LegendPos = leg?.Element(c + "legendPos")?.Attribute("val")?.Value;
                    info.DataLabels = plotArea?.Descendants(c + "dLbls").Any() == true;

                    // Series
                    foreach (var ser in plotArea?.Descendants().Where(e => e.Name.LocalName == "ser") ?? Enumerable.Empty<XElement>())
                    {
                        var si = new AutoSeriesInfo();
                        var tx = ser.Element(c + "tx");
                        (si.Name, si.NameRef) = ReadChartTextAuto(tx, c, a);

                        var cat = ser.Element(c + "cat");
                        si.CatRef = cat?.Element(c + "strRef")?.Element(c + "f")?.Value
                                    ?? cat?.Element(c + "numRef")?.Element(c + "f")?.Value;

                        var val = ser.Element(c + "val");
                        si.ValRef = val?.Element(c + "numRef")?.Element(c + "f")?.Value
                                    ?? val?.Element(c + "strRef")?.Element(c + "f")?.Value;

                        info.Series.Add(si);
                    }

                    if (!result.TryGetValue(sheetName, out var list)) result[sheetName] = list = new();
                    list.Add(info);
                }
            }
        }

        return result;

        // ---- helpers ----
        static (string? txt, string? cellRef) ReadChartTextAuto(XElement? node, XNamespace cns, XNamespace ans)
        {
            if (node == null) return (null, null);
            var tx = node.Element(cns + "tx");
            if (tx == null) return (null, null);

            var rich = tx.Element(cns + "rich");
            if (rich != null)
            {
                var text = string.Join("", rich.Descendants(ans + "t").Select(t => t.Value));
                return (text, null);
            }
            var strRef = tx.Element(cns + "strRef");
            var f = strRef?.Element(cns + "f")?.Value;
            return (null, f);
        }
    }

    private static string DetectChartType(System.Xml.Linq.XElement? plotArea, System.Xml.Linq.XNamespace c)
    {
        if (plotArea == null) return "";
        if (plotArea.Element(c + "bar3DChart") != null)
            return string.Equals(plotArea.Element(c + "bar3DChart")?.Element(c + "barDir")?.Attribute("val")?.Value, "col", StringComparison.OrdinalIgnoreCase) ? "column3D" : "bar3D";
        var bar = plotArea.Element(c + "barChart");
        if (bar != null)
            return string.Equals(bar.Element(c + "barDir")?.Attribute("val")?.Value, "col", StringComparison.OrdinalIgnoreCase) ? "column" : "bar";
        if (plotArea.Element(c + "line3DChart") != null) return "line3D";
        if (plotArea.Element(c + "area3DChart") != null) return "area3D";
        if (plotArea.Element(c + "pie3DChart") != null) return "pie3D";
        if (plotArea.Element(c + "pieChart") != null) return "pie";
        if (plotArea.Element(c + "areaChart") != null) return "area";
        if (plotArea.Element(c + "lineChart") != null) return "line";
        if (plotArea.Element(c + "scatterChart") != null) return "scatter";
        if (plotArea.Element(c + "bubbleChart") != null) return "bubble";
        if (plotArea.Element(c + "radarChart") != null) return "radar";
        if (plotArea.Element(c + "stockChart") != null) return "stock";
        return "";
    }



}
