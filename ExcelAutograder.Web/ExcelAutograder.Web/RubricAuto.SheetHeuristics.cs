using System.Text.Json;
using ClosedXML.Excel;

public static partial class RubricAuto
{
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
        var sf = NormalizeFormulaAuto(summaryCell.FormulaA1);

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
            var f = NormalizeFormulaAuto(c.FormulaA1);
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

    /// <summary>
    /// Build pivot-layout rules using ClosedXML (or your existing approach) for a single sheet.
    /// </summary>
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

    /// <summary>
    /// If your ZIP scan didn’t find pivots, you may have a heuristic fallback to
    /// indicate pivot presence/shape; keep that here.
    /// </summary>
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

    private static List<Rule> BuildChecksForRanges(
        IXLWorksheet ws,
        IEnumerable<IXLRangeAddress> ranges,
        string? sectionName = null)
    {
        var checks = new List<Rule>();
        var seenCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // --- 1) PER-CELL rules (formula OR value) for every address in each range ---
        foreach (var ra in ranges)
        {
            var rect = ws.Range(ra);
            int r1 = rect.RangeAddress.FirstAddress.RowNumber;
            int c1 = rect.RangeAddress.FirstAddress.ColumnNumber;
            int r2 = rect.RangeAddress.LastAddress.RowNumber;
            int c2 = rect.RangeAddress.LastAddress.ColumnNumber;

            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                {
                    var cell = ws.Cell(r, c);
                    var a1 = cell.Address.ToStringRelative();
                    if (!seenCells.Add(a1)) continue;

                    var f = NormalizeFormulaAuto(cell.FormulaA1);
                    if (!string.IsNullOrWhiteSpace(f))
                    {
                        // formula rule
                        var rule = new Rule
                        {
                            Type = "formula",
                            Points = 0.5,
                            Cell = a1,
                            Note = "Formula from key (selected range)",
                            ExpectedFromKey = true,
                            ExpectedFormula = f,
                            Expected = System.Text.Json.JsonSerializer.SerializeToElement(CleanExpected(KeyCellExpectedText(cell))),
                            Section = sectionName
                        };
                        if (HasAbsoluteRef(f)) rule.RequireAbsolute = true;
                        checks.Add(rule);
                    }
                    else
                    {
                        // value rule (including blanks)
                        checks.Add(new Rule
                        {
                            Type = "value",
                            Points = 0.5,
                            Cell = a1,
                            Note = "Value from key (selected range)",
                            ExpectedFromKey = true,
                            Expected = System.Text.Json.JsonSerializer.SerializeToElement(
                                CleanExpected(KeyCellExpectedText(cell) ?? string.Empty)),
                            Section = sectionName
                        });
                    }
                }
        }

        // --- 2) Optional: number-format per range when there is a consistent pattern ---
        foreach (var ra in ranges)
        {
            var r = ws.Range(ra);
            var fmt = GetRangeNumberFormat(ws, r.RangeAddress.ToStringRelative());
            if (string.IsNullOrWhiteSpace(fmt)) continue;

            checks.Add(new Rule
            {
                Type = "range_format",
                Points = 0.25,
                Range = r.RangeAddress.ToStringRelative(),
                Note = "Number format (selected range)",
                Format = new FormatSpec { NumberFormat = fmt },
                Section = sectionName
            });
        }

        // --- 3) Summary row immediately below each range (keep as-is) ---
        foreach (var ra in ranges)
        {
            var r = ws.Range(ra);
            var lastRow = r.RangeAddress.LastAddress.RowNumber;
            var firstCol = r.RangeAddress.FirstAddress.ColumnNumber;
            var lastCol = r.RangeAddress.LastAddress.ColumnNumber;

            for (int col = firstCol; col <= lastCol; col++)
            {
                var cell = ws.Cell(lastRow + 1, col);
                var f = NormalizeFormulaAuto(cell.FormulaA1);
                if (string.IsNullOrWhiteSpace(f)) continue;

                checks.Add(new Rule
                {
                    Type = "formula",
                    Points = 1.0,
                    Cell = cell.Address.ToString(),
                    Note = "Summary formula (selected range)",
                    ExpectedFromKey = true,
                    ExpectedFormula = f,
                    Expected = System.Text.Json.JsonSerializer.SerializeToElement(CleanExpected(KeyCellExpectedText(cell))),
                    Section = sectionName
                });
            }
        }

        // --- TABLE rules for any table intersecting the selected ranges ---
        {
            // Build a fast list of range objects for intersection checks
            var selRects = ranges
                .Select(ra => ws.Range(ra))
                .ToList();

            foreach (var tbl in ws.Tables)
            {
                var tRange = ws.Range(tbl.RangeAddress);

                // include table only if it touches any selected range
                bool intersects = selRects.Any(r => r.Intersects(tRange));
                if (!intersects) continue;

                // header names
                var cols = tbl.Fields
                    .Select(f => (f.Name ?? string.Empty).Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();
                if (cols.Count == 0) continue;

                // full A1 (header + body)
                var fullRangeA1 = tbl.RangeAddress.ToStringRelative();

                // body dims + optional sample (safe)
                var body = tbl.DataRange;
                int bodyRows = body?.RowCount() ?? 0;
                int bodyCols = body?.ColumnCount() ?? 0;

                // Avoid duplicates if a same table rule already added
                bool already =
                    checks.Any(c => c.Type == "table" &&
                                    (c.Table?.NameLike ?? "").Equals(tbl.Name, StringComparison.OrdinalIgnoreCase));

                if (!already)
                {
                    checks.Add(new Rule
                    {
                        Type = "table",
                        Points = 1.0,
                        Section = sectionName, // preserve section label if provided
                        Note = $"Table '{tbl.Name}' columns",
                        Table = new TableSpec
                        {
                            Sheet = ws.Name,
                            NameLike = tbl.Name,

                            // columns (order not required by default)
                            Columns = cols,
                            RequireOrder = false,

                            // dimensions + location (helps grader verify)
                            RangeRef = fullRangeA1,
                            Rows = bodyRows,
                            Cols = bodyCols,

                            // students can have extra rows/cols as they add data
                            AllowExtraRows = true,
                            AllowExtraCols = true
                        }
                    });
                }
            }
        }

        // --- PIVOT rules on this sheet (optional: keep only those that touch selected ranges) ---
        {
            // (A) if you want to keep *all* pivots on this sheet in this section:
            var pivotsObj = ws.GetType().GetProperty("PivotTables")?.GetValue(ws);
            foreach (var pt in AsEnum(pivotsObj))
            {
                var ptName = GetStrProp(pt, "Name") ?? ws.Name + " Pivot";

                // Collect layout, same as BuildPivotRulesForSheet(...)
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

                if (rows.Count == 0 && cols.Count == 0 && filters.Count == 0 && values.Count == 0)
                    continue;

                checks.Add(new Rule
                {
                    Type = "pivot_layout",
                    Points = 1.5,
                    Section = sectionName,                  // put it under the user’s section
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
        }


        return checks;
    }
}

