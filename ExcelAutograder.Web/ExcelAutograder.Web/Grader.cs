// Grader.cs
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;

public static class Grader
{
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
                // mark all checks for that sheet as failed
                foreach (var rule in spec.Checks)
                {
                    var id = rule.Cell ?? rule.Range ?? sheetName;
                    results.Add(new CheckResult($"{rule.Type}:{id}", rule.Points, 0, false,
                        $"Sheet '{sheetName}' missing"));
                }
                continue;
            }

            var wsK = FindWorksheet(wbKey, sheetName);

            foreach (var rule in spec.Checks)
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

    public static object Run(XLWorkbook wbKey, XLWorkbook wbStudent, Rubric rubric, byte[]? studentZipBytes)
    {
        _studentZipBytes = studentZipBytes;      // store for conditional-format grader
        try { return Run(wbKey, wbStudent, rubric); }
        finally { _studentZipBytes = null; }
    }

    // backing field
    private static byte[]? _studentZipBytes;


    // ---- Router
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

    // =====================
    // VALUE / FORMULA
    // =====================

    private static CheckResult GradeValue(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var cellAddr = rule.Cell ?? throw new Exception("value check missing 'cell'");
        var pts = rule.Points;
        var tol = rule.Tolerance ?? 0.0;

        (bool ok, string reason) OneOption(RuleOption opt)
        {
            object? expected;
            if (rule.ExpectedFromKey == true)
            {
                expected = wsK?.Cell(cellAddr).Value;
            }
            else if (opt.ExpectedRegex is not null)
            {
                var sval = Normalize(wsS.Cell(cellAddr).Value);
                bool match = Regex.IsMatch(sval, $"^{opt.ExpectedRegex}$");
                return (match, $"value='{sval}' regex='{opt.ExpectedRegex}'");
            }
            else if (opt.Expected.HasValue)
            {
                expected = JsonToNet(opt.Expected.Value);
            }
            else if (rule.Expected.HasValue)
            {
                expected = JsonToNet(rule.Expected.Value);
            }
            else
            {
                return (false, "No expected value provided.");
            }

            var sVal = wsS.Cell(cellAddr).Value;
            if (TryToDouble(expected, out var ed) && TryToDouble(sVal, out var sd))
            {
                bool match = Math.Abs(sd - ed) <= tol;
                return (match, $"value={sd} expected={ed} tol={tol}");
            }
            else
            {
                var actualStr = sVal.ToString()?.Trim() ?? "";
                var expectedStr = (expected?.ToString() ?? "").Trim();

                bool caseSensitive = opt.CaseSensitive ?? rule.CaseSensitive ?? false;
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                bool eq = string.Equals(actualStr, expectedStr, comparison);
                return (eq, $"value='{actualStr}' expected='{expectedStr}' (case {(caseSensitive ? "sensitive" : "insensitive")})");
            }
        }

        var result = rule.AnyOf is { Count: > 0 }
            ? AnyOfMatch(rule.AnyOf, OneOption)
            : OneOption(new RuleOption());

        return new CheckResult($"value:{cellAddr}", pts, result.ok ? pts : 0, result.ok, result.reason);
    }

    private static CheckResult GradeFormula(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var cellAddr = rule.Cell ?? throw new Exception("formula check missing 'cell'");
        var pts = rule.Points;

        // student's formula (raw + normalized)
        string sRaw = wsS.Cell(cellAddr).FormulaA1 ?? "";
        string sNorm = NormalizeFormula(sRaw);

        // Build options to try (any_of overrides rule-level)
        var opts = (rule.AnyOf is { Count: > 0 }
            ? rule.AnyOf.Select(o => (o.ExpectedFormula, o.ExpectedFormulaRegex, o.ExpectedFromKey, o.RequireAbsolute, "option"))
            : new[] { (rule.ExpectedFormula, rule.ExpectedFormulaRegex, rule.ExpectedFromKey, rule.RequireAbsolute, "rule") }
        ).ToList();

        var reasons = new List<string>();

        foreach (var (litExp, rxExp, fromKey, requireAbs, origin) in opts)
        {
            // 1) Literal expected
            var litNorm = NormalizeFormula(litExp ?? "");
            if (!string.IsNullOrEmpty(litNorm))
            {
                bool contentOk = sNorm == litNorm;

                if (contentOk)
                {
                    if (requireAbs == true)
                    {
                        var missing = MissingAbsolutesFromExpected(litExp!, sRaw);
                        if (missing.Count > 0)
                            return new CheckResult(
                                $"formula:{cellAddr}", pts, 0, false,
                                $"formula correct but missing required absolutes: {string.Join(", ", missing)} ({origin}); got='{sRaw}' expected='{litExp}'"
                            );
                    }

                    return new CheckResult(
                        $"formula:{cellAddr}", pts, pts, true,
                        $"formula='{sNorm}' expected='{litNorm}' ({origin})"
                    );
                }

                // content mismatch; still show absolute problems for context if requested
                if (requireAbs == true)
                {
                    var missing = MissingAbsolutesFromExpected(litExp!, sRaw);
                    if (missing.Count > 0) reasons.Add($"missing required absolutes: {string.Join(", ", missing)}");
                }

                reasons.Add($"formula='{sNorm}' expected='{litNorm}' ({origin}); got='{sRaw}'");
                continue;
            }

            // 2) Regex expected
            if (!string.IsNullOrWhiteSpace(rxExp))
            {
                bool contentOk = Regex.IsMatch(sNorm, $"^{rxExp}$", RegexOptions.IgnoreCase);

                if (contentOk)
                {
                    if (requireAbs == true)
                    {
                        // No concrete expected text → fall back to "all must be absolute"
                        var absInfo = MissingAbsoluteRefs(sRaw);
                        if (!absInfo.ok)
                            return new CheckResult(
                                $"formula:{cellAddr}", pts, 0, false,
                                $"formula correct (regex) but missing absolutes: {string.Join(", ", absInfo.missing)} ({origin}); got='{sRaw}'"
                            );
                    }

                    return new CheckResult(
                        $"formula:{cellAddr}", pts, pts, true,
                        $"formula='{sNorm}' regex='{rxExp}' ({origin})"
                    );
                }

                reasons.Add($"formula='{sNorm}' regex='{rxExp}' no match ({origin}); got='{sRaw}'");
                continue;
            }

            // 3) Expected from key
            if (fromKey == true)
            {
                if (wsK is null) { reasons.Add($"key sheet missing ({origin})"); continue; }

                var kc = wsK.Cell(cellAddr);
                var kRaw = kc.HasFormula ? (kc.FormulaA1 ?? kc.FormulaR1C1 ?? "") : "";
                var kNorm = NormalizeFormula(kRaw);

                if (!string.IsNullOrEmpty(kNorm) && sNorm == kNorm)
                {
                    if (requireAbs == true)
                    {
                        var missing = MissingAbsolutesFromExpected(kRaw, sRaw);
                        if (missing.Count > 0)
                            return new CheckResult(
                                $"formula:{cellAddr}", pts, 0, false,
                                $"formula matches key but missing required absolutes: {string.Join(", ", missing)} (from key); got='{sRaw}' expected='{kRaw}'"
                            );
                    }

                    return new CheckResult(
                        $"formula:{cellAddr}", pts, pts, true,
                        $"formula='{sNorm}' expected='{kNorm}' (from key)"
                    );
                }

                if (requireAbs == true)
                {
                    var missing = MissingAbsolutesFromExpected(kRaw, sRaw);
                    if (missing.Count > 0) reasons.Add($"missing required absolutes: {string.Join(", ", missing)}");
                }

                reasons.Add(kc.HasFormula
                    ? $"formula='{sNorm}' expected='{kNorm}' (from key); got='{sRaw}'"
                    : "key cell has no formula");
                continue;
            }

            // No expectation configured
            reasons.Add($"no expected provided ({origin})");
        }

        // Content did not match any option; return accumulated reasons
        var failMsg = string.Join(" | ", reasons);
        return new CheckResult($"formula:{cellAddr}", pts, 0, false, failMsg);
    }





    // Helper (unchanged): detect absolute refs
    private static (bool any, bool col, bool row) InspectAbsoluteRefs(string? formulaA1)
    {
        var f = formulaA1 ?? "";
        bool col = Regex.IsMatch(f, @"\$[A-Za-z]");
        bool row = Regex.IsMatch(f, @"\$\d");
        return (col || row, col, row);
    }

    // Returns ok + list of refs that are NOT fully absolute ($col$row)
    private static (bool ok, List<string> missing) MissingAbsoluteRefs(string? formulaA1)
    {
        var text = formulaA1 ?? string.Empty;
        // Rough A1 finder: optional 'Sheet'!, then $A$1/$A1/A$1/A1. Skips functions, etc.
        var rx = new Regex(@"(?<![A-Z0-9_])(?:'[^']+'!)?(\$?)[A-Za-z]{1,3}(\$?)[0-9]+",
                           RegexOptions.IgnoreCase);
        var missing = new List<string>();

        foreach (Match m in rx.Matches(text))
        {
            bool colAbs = m.Groups[1].Value == "$";
            bool rowAbs = m.Groups[2].Value == "$";
            if (!(colAbs && rowAbs))
            {
                var token = m.Value;
                int bang = token.LastIndexOf('!');
                if (bang >= 0 && bang + 1 < token.Length) token = token[(bang + 1)..]; // strip sheet
                missing.Add(token);
            }
        }
        return (missing.Count == 0, missing);
    }

    // Extract single-cell endpoints in order (e.g., "A1:B10" -> ["A1","B10"]).
    private static List<(bool colAbs, bool rowAbs, string token)> ExtractEndpoints(string? formulaA1)
    {
        var res = new List<(bool colAbs, bool rowAbs, string token)>();
        if (string.IsNullOrWhiteSpace(formulaA1)) return res;

        var rx = new Regex(@"(?<![A-Z0-9_])(?:'[^']+'!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)", RegexOptions.IgnoreCase);
        foreach (Match m in rx.Matches(formulaA1))
        {
            bool colAbs = m.Groups[1].Value == "$";
            bool rowAbs = m.Groups[3].Value == "$";
            // strip sheet if present for nicer messages
            string token = m.Value;
            int bang = token.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < token.Length) token = token[(bang + 1)..];
            res.Add((colAbs, rowAbs, token));
        }
        return res;
    }

    // Compare absolutes against the EXPECTED formula.
    // Returns the list of endpoints that should be absolute (per expected)
    // but are not absolute in the student's formula at the same position.
    private static List<string> MissingAbsolutesFromExpected(string expectedA1, string studentA1)
    {
        var exp = ExtractEndpoints(expectedA1);
        var got = ExtractEndpoints(studentA1);

        var missing = new List<string>();
        int n = Math.Min(exp.Count, got.Count);

        for (int i = 0; i < n; i++)
        {
            var e = exp[i]; var g = got[i];
            if (e.colAbs && !g.colAbs) missing.Add(g.token);
            if (e.rowAbs && !g.rowAbs) missing.Add(g.token);
        }

        // If student has fewer endpoints than expected, any remaining absolutes are missing.
        for (int i = got.Count; i < exp.Count; i++)
            if (exp[i].colAbs || exp[i].rowAbs) missing.Add(exp[i].token);

        // dedupe
        return missing.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }


    // =====================
    // FORMAT
    // =====================

    private static CheckResult GradeFormat(Rule rule, IXLWorksheet wsS)
    {
        var cellAddr = rule.Cell ?? throw new Exception("format check missing 'cell'");
        var pts = rule.Points;
        var cell = wsS.Cell(cellAddr);

        (bool ok, string reason) OneFormat(RuleOption opt) =>
            FormatMatches(cell, opt.Format ?? rule.Format ?? new());

        var result = rule.AnyOf is { Count: > 0 }
            ? AnyOfMatch(rule.AnyOf, OneFormat)
            : OneFormat(new RuleOption());

        return new CheckResult($"format:{cellAddr}", pts, result.ok ? pts : 0, result.ok, result.reason);
    }

    private static (bool ok, string reason) FormatMatches(IXLCell c, FormatSpec fmt)
    {
        var reasons = new List<string>();
        var style = c.Style;

        // font group
        if (fmt.Font is not null)
        {
            if (fmt.Font.Name is not null && style.Font.FontName != fmt.Font.Name) reasons.Add("font name");
            if (fmt.Font.Size is not null && Math.Abs(style.Font.FontSize - fmt.Font.Size.Value) > 0.01) reasons.Add("font size");
            if (fmt.Font.Bold is not null && style.Font.Bold != fmt.Font.Bold.Value) reasons.Add("font bold");
            if (fmt.Font.Italic is not null && style.Font.Italic != fmt.Font.Italic.Value) reasons.Add("font italic");
        }
        if (fmt.Bold is not null && style.Font.Bold != fmt.Bold.Value) reasons.Add("bold");
        if (fmt.Italic is not null && style.Font.Italic != fmt.Italic.Value) reasons.Add("italic");

        if (fmt.NumberFormat is not null && (style.NumberFormat.Format ?? "") != fmt.NumberFormat) reasons.Add("number_format");

        if (fmt.Fill?.Rgb is not null)
        {
            var want = NormalizeArgb(fmt.Fill.Rgb);
            var got = XLColorToArgb(style.Fill.BackgroundColor);
            if (!ArgbEqual(got, want)) reasons.Add($"fill ({got} != {want})");
        }

        if (fmt.Alignment is not null)
        {
            if (fmt.Alignment.Horizontal is not null &&
                !string.Equals(style.Alignment.Horizontal.ToString(), fmt.Alignment.Horizontal, StringComparison.OrdinalIgnoreCase))
                reasons.Add("alignment horizontal");

            if (fmt.Alignment.Vertical is not null &&
                !string.Equals(style.Alignment.Vertical.ToString(), fmt.Alignment.Vertical, StringComparison.OrdinalIgnoreCase))
                reasons.Add("alignment vertical");
        }

        if (fmt.Border?.Outline is not null)
        {
            bool outlined =
                style.Border.LeftBorder != XLBorderStyleValues.None ||
                style.Border.RightBorder != XLBorderStyleValues.None ||
                style.Border.TopBorder != XLBorderStyleValues.None ||
                style.Border.BottomBorder != XLBorderStyleValues.None;

            if (outlined != fmt.Border.Outline.Value) reasons.Add("border outline");
        }

        bool ok = reasons.Count == 0;
        return (ok, ok ? "format ok" : string.Join(", ", reasons));
    }

    // =====================
    // RANGE CHECKS
    // =====================

    private static CheckResult GradeRangeValue(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var refRange = rule.Range ?? throw new Exception("range_value check missing 'range'");
        var pts = rule.Points;
        var tol = rule.Tolerance ?? 0.0;

        var rangeS = wsS.Range(refRange);
        int total = 0, correct = 0;

        foreach (var cS in rangeS.CellsUsed(XLCellsUsedOptions.All))
        {
            total++;
            object? expected;
            if (rule.ExpectedFromKey == true && wsK is not null)
                expected = wsK.Cell(cS.Address.ToString()).Value;
            else if (rule.Expected.HasValue)
                expected = JsonToNet(rule.Expected.Value);
            else
                expected = null;

            var sVal = cS.Value;
            bool ok;
            if (TryToDouble(expected, out var ed) && TryToDouble(sVal, out var sd))
                ok = Math.Abs(sd - ed) <= tol;
            else
                ok = Normalize(sVal) == Normalize(expected);

            if (ok) correct++;
        }

        double frac = total == 0 ? 0 : (double)correct / total;
        double earned = pts * frac;

        return new CheckResult($"range_value:{refRange}", pts, earned, Math.Abs(frac - 1.0) < 1e-9, $"{correct}/{total} cells matched");
    }

    private static CheckResult GradeRangeFormula(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var refRange = rule.Range ?? throw new Exception("range_formula check missing 'range'");
        var pts = rule.Points;

        var rangeS = wsS.Range(refRange);
        var regex = rule.ExpectedFormulaRegex is not null ? new Regex($"^{rule.ExpectedFormulaRegex}$") : null;
        var expectedLiteral = NormalizeFormula(rule.ExpectedFormula ?? "");

        int total = 0, correct = 0;

        foreach (var cS in rangeS.CellsUsed(XLCellsUsedOptions.All))
        {
            total++;

            var sF = NormalizeFormula(cS.FormulaA1);
            bool ok;

            if (rule.ExpectedFromKey == true && wsK is not null)
            {
                var kF = NormalizeFormula(wsK.Cell(cS.Address.ToString()).FormulaA1);
                ok = sF == kF;
            }
            else if (regex is not null)
            {
                ok = regex.IsMatch(sF);
            }
            else
            {
                ok = sF == expectedLiteral;
            }

            if (ok) correct++;
        }

        double frac = total == 0 ? 0 : (double)correct / total;
        double earned = pts * frac;

        return new CheckResult($"range_formula:{refRange}", pts, earned, Math.Abs(frac - 1.0) < 1e-9, $"{correct}/{total} formulas matched");
    }

    private static CheckResult GradeRangeFormat(Rule rule, IXLWorksheet wsS)
    {
        var refRange = rule.Range ?? throw new Exception("range_format check missing 'range'");
        var pts = rule.Points;
        var rangeS = wsS.Range(refRange);

        int total = 0, correct = 0;

        foreach (var c in rangeS.CellsUsed(XLCellsUsedOptions.All))
        {
            total++;

            bool OkOne(IXLCell cell, RuleOption opt) =>
                FormatMatches(cell, opt.Format ?? rule.Format ?? new()).ok;

            bool ok = rule.AnyOf is { Count: > 0 }
                ? rule.AnyOf.Any(opt => OkOne(c, opt))
                : OkOne(c, new RuleOption());

            if (ok) correct++;
        }

        double frac = total == 0 ? 0 : (double)correct / total;
        double earned = pts * frac;

        return new CheckResult($"range_format:{refRange}", pts, earned, Math.Abs(frac - 1.0) < 1e-9, $"{correct}/{total} cells match formatting");
    }

    private static CheckResult GradeRangeSequence(Rule rule, IXLWorksheet wsS)
    {
        var refRange = rule.Range ?? throw new Exception("range_sequence check missing 'range'");
        var pts = rule.Points;
        var start = rule.Start ?? 1.0;
        var step = rule.Step ?? 1.0;

        var range = wsS.Range(refRange);
        int total = 0, correct = 0;
        double expected = start;

        foreach (var c in range.Cells()) // include blanks
        {
            total++;
            if (TryToDouble(c.Value, out var dv) && Math.Abs(dv - expected) < 1e-9)
                correct++;
            expected += step;
        }

        double frac = total == 0 ? 0 : (double)correct / total;
        return new CheckResult($"range_sequence:{refRange}", pts, pts * frac,
            Math.Abs(frac - 1.0) < 1e-9, $"{correct}/{total} cells match sequence");
    }

    private static CheckResult GradeRangeNumeric(Rule rule, IXLWorksheet wsS)
    {
        var refRange = rule.Range ?? throw new Exception("range_numeric check missing 'range'");
        var pts = rule.Points;
        var range = wsS.Range(refRange);

        int total = 0, correct = 0;

        foreach (var c in range.Cells())
        {
            var text = c.GetString();
            if (string.IsNullOrWhiteSpace(text)) continue; // ignore blanks
            total++;
            if (TryToDouble(c.Value, out _)) correct++;
        }

        // if all blanks, consider it 0/pts (or change to full credit if you prefer)
        if (total == 0) return new CheckResult($"range_numeric:{refRange}", pts, 0, false, "no non-blank cells");

        double frac = (double)correct / total;
        return new CheckResult($"range_numeric:{refRange}", pts, pts * frac,
            Math.Abs(frac - 1.0) < 1e-9, $"{correct}/{total} numeric (blanks ignored)");
    }

    // =====================
    // CHART TABLE
    // =====================
    private class ChartInfo
    {
        public string Sheet = "";
        public string Name = "";                // "Chart 1", etc.
        public string Type = "";                // normalized: line/column/bar/pie/scatter/area/doughnut/radar/bubble
        public string? Title, TitleRef;
        public string? LegendPos;
        public bool DataLabels;
        public string? XTitle, YTitle;
        public List<SeriesInfo> Series = new();
    }
    private class SeriesInfo
    {
        public string? Name, NameRef, CatRef, ValRef;
    }

    private static CheckResult GradeChart(Rule rule, XLWorkbook wbS)
    {
        var pts = rule.Points;
        var spec = rule.Chart;
        if (spec is null)
            return new CheckResult("chart", pts, 0, false, "No chart spec provided");

        if (_studentZipBytes is null)
            return new CheckResult("chart", pts, 0, false, "Chart checks require workbook bytes");

        var bySheet = ParseChartsFromZip(_studentZipBytes);  // sheet -> charts
        IEnumerable<ChartInfo> pool = bySheet.SelectMany(kv => kv.Value);
        if (!string.IsNullOrWhiteSpace(spec.Sheet))
            pool = pool.Where(c => string.Equals(c.Sheet, spec.Sheet, StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrWhiteSpace(spec.NameLike))
            pool = pool.Where(c => c.Name.IndexOf(spec.NameLike!, StringComparison.OrdinalIgnoreCase) >= 0);

        // Scoring: each declared attribute is a "check".
        int checks = 0, hits = 0;
        var notes = new List<string>();

        ChartInfo? best = null;
        double bestScore = -1;

        foreach (var ch in pool)
        {
            int cks = 0, ok = 0;
            void req(bool cond, string okMsg, string badMsg)
            { cks++; if (cond) { ok++; notes.Add($"{ch.Name}: {okMsg}"); } else notes.Add($"{ch.Name}: {badMsg}"); }

            // type
            if (!string.IsNullOrWhiteSpace(spec.Type))
                req(string.Equals(ch.Type, spec.Type, StringComparison.OrdinalIgnoreCase),
                    $"type={ch.Type}", $"type got {ch.Type} expected {spec.Type}");

            // title
            if (!string.IsNullOrWhiteSpace(spec.Title))
                req(string.Equals((ch.Title ?? ""), spec.Title, StringComparison.Ordinal),
                    $"title='{ch.Title}'", $"title got '{ch.Title ?? ""}' expected '{spec.Title}'");
            if (!string.IsNullOrWhiteSpace(spec.TitleRef))
                req(string.Equals((ch.TitleRef ?? ""), spec.TitleRef, StringComparison.OrdinalIgnoreCase),
                    $"title_ref={ch.TitleRef}", $"title_ref got {ch.TitleRef ?? ""} expected {spec.TitleRef}");

            // legend
            if (!string.IsNullOrWhiteSpace(spec.LegendPos))
                req(string.Equals((ch.LegendPos ?? ""), spec.LegendPos, StringComparison.OrdinalIgnoreCase),
                    $"legendPos={ch.LegendPos}", $"legendPos got {ch.LegendPos ?? ""} expected {spec.LegendPos}");

            // data labels
            if (spec.DataLabels.HasValue)
                req(ch.DataLabels == spec.DataLabels.Value,
                    $"dataLabels={ch.DataLabels}", $"dataLabels got {ch.DataLabels} expected {spec.DataLabels}");

            // axis titles
            if (!string.IsNullOrWhiteSpace(spec.XTitle))
                req(string.Equals((ch.XTitle ?? ""), spec.XTitle, StringComparison.Ordinal),
                    $"xTitle='{ch.XTitle}'", $"xTitle got '{ch.XTitle ?? ""}' expected '{spec.XTitle}'");
            if (!string.IsNullOrWhiteSpace(spec.YTitle))
                req(string.Equals((ch.YTitle ?? ""), spec.YTitle, StringComparison.Ordinal),
                    $"yTitle='{ch.YTitle}'", $"yTitle got '{ch.YTitle ?? ""}' expected '{spec.YTitle}'");

            // series: every specified expected series must be present (match by CatRef+ValRef, optionally name/nameRef)
            if (spec.Series is { Count: > 0 })
            {
                foreach (var exp in spec.Series)
                {
                    cks++;
                    bool found = ch.Series.Any(s =>
                        (string.IsNullOrWhiteSpace(exp.CatRef) || string.Equals((s.CatRef ?? ""), exp.CatRef, StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrWhiteSpace(exp.ValRef) || string.Equals((s.ValRef ?? ""), exp.ValRef, StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrWhiteSpace(exp.Name) || string.Equals((s.Name ?? ""), exp.Name, StringComparison.Ordinal)) &&
                        (string.IsNullOrWhiteSpace(exp.NameRef) || string.Equals((s.NameRef ?? ""), exp.NameRef, StringComparison.OrdinalIgnoreCase))
                    );
                    if (found) { ok++; notes.Add($"{ch.Name}: series ok ({exp.CatRef} / {exp.ValRef})"); }
                    else notes.Add($"{ch.Name}: series missing ({exp.CatRef} / {exp.ValRef})");
                }
            }

            // choose best chart (most hits)
            double score = (cks == 0) ? 0 : (double)ok / cks;
            if (score > bestScore) { bestScore = score; best = ch; checks = cks; hits = ok; }
        }

        if (best == null)
            return new CheckResult("chart", pts, 0, false, "No charts matched the selection pool");

        double earned = (checks == 0) ? 0 : pts * (double)hits / checks;
        bool pass = hits == checks;
        return new CheckResult($"chart:{best.Sheet}/{best.Name}", pts, earned, pass,
            $"matched {hits}/{checks} checks; type={best.Type}; title='{best.Title ?? ""}'");
    }

    // Parse charts via OOXML
    private static Dictionary<string, List<ChartInfo>> ParseChartsFromZip(byte[] zipBytes)
    {
        using var ms = new MemoryStream(zipBytes);
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
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

        // Map sheet index → name (1-based like sheet1.xml, …)
        var wb = XDocument.Parse(ReadEntryText("xl/workbook.xml"));
        var sheetIndexToName = new Dictionary<int, string>();
        int idx = 1;
        var sheetsEl = wb.Root?.Element(nsMain + "sheets");
        if (sheetsEl != null)
            foreach (var sh in sheetsEl.Elements(nsMain + "sheet"))
                sheetIndexToName[idx++] = (string?)sh.Attribute("name") ?? $"Sheet{idx - 1}";

        var result = new Dictionary<string, List<ChartInfo>>(StringComparer.OrdinalIgnoreCase);

        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsTxt = ReadEntryText(relsPath);
            if (string.IsNullOrEmpty(relsTxt)) continue;

            var relsDoc = XDocument.Parse(relsTxt);
            var drawingTargets = relsDoc.Root?
                 .Elements(pkg + "Relationship")
                 .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/drawing") == true)
                 .Select(r => ((string?)r.Attribute("Target"))?.TrimStart('/').Replace("../", "xl/"))
                 .Where(t => !string.IsNullOrWhiteSpace(t))
                 .ToList() ?? new List<string>();

            foreach (var drawingRelTarget in drawingTargets)
            {
                // Normalize path like "../drawings/drawing1.xml"
                var drPath = drawingRelTarget!.StartsWith("xl/") ? drawingRelTarget! : $"xl/{drawingRelTarget}";
                var drXmlTxt = ReadEntryText(drPath);

                if (string.IsNullOrEmpty(drXmlTxt)) continue;

                var drXml = XDocument.Parse(drXmlTxt);
                // Build a map: r:id -> frame name ("Chart 1", "Chart 2", …)
                var frameMap = drXml.Descendants(xdr + "graphicFrame")
                    .Select(gf => new {
                        Name = gf.Element(xdr + "nvGraphicFramePr")?.Element(xdr + "cNvPr")?.Attribute("name")?.Value,
                        Rid = gf.Descendants(a + "graphicData").Descendants(c + "chart")
                                 .FirstOrDefault()?.Attribute(rel + "id")?.Value
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.Rid))
                    .ToDictionary(x => x.Rid!, x => x.Name ?? "Chart", StringComparer.OrdinalIgnoreCase);

                string drRelsPath = drPath.Replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
                var drRelsTxt = ReadEntryText(drRelsPath);
                if (string.IsNullOrEmpty(drRelsTxt)) continue;
                var drRels = XDocument.Parse(drRelsTxt);

                // For each <xdr:graphicFrame> → <c:chart r:id="...">
                var chartIds = drXml.Descendants(xdr + "graphicFrame")
                    .Select(gf => gf.Descendants(a + "graphicData").Descendants(c + "chart").FirstOrDefault())
                    .Where(ch => ch != null)
                    .Select(ch => (string?)ch!.Attribute(rel + "id"))
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .ToList();

                // Resolve r:id → charts/chartN.xml
                foreach (var rid in chartIds)
                {
                    var target = drRels.Root?
                        .Elements(pkg + "Relationship")
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;

                    if (string.IsNullOrWhiteSpace(target)) continue;

                    var tgt = (target ?? "").Replace("\\", "/");
                    var chartPath = tgt.StartsWith("/") ? tgt.TrimStart('/')
                                 : tgt.StartsWith("../") ? "xl/" + tgt.Substring(3)
                                 : tgt.StartsWith("xl/") ? tgt : "xl/" + tgt;

                    var chXmlTxt = ReadEntryText(chartPath);

                    if (string.IsNullOrEmpty(chXmlTxt)) continue;

                    var chXml = XDocument.Parse(chXmlTxt);
                    var chart = new ChartInfo { Sheet = sheetName };

                    // correct name per rid
                    chart.Name = frameMap.TryGetValue(rid!, out var nm) ? nm : "Chart";

                    // detect type, then read title/legend/series...
                    var plotArea = chXml.Descendants(c + "plotArea").FirstOrDefault();
                    chart.Type = DetectChartType(plotArea);

                    // Title
                    var titleEl = chXml.Descendants(c + "title").FirstOrDefault();
                    (chart.Title, chart.TitleRef) = ReadChartText(titleEl, c, a);

                    // Axis titles (cat/val axes)
                    var catAx = plotArea?.Elements(c + "catAx").FirstOrDefault();
                    var valAx = plotArea?.Elements(c + "valAx").FirstOrDefault();
                    (chart.XTitle, _) = ReadChartText(catAx?.Element(c + "title"), c, a);
                    (chart.YTitle, _) = ReadChartText(valAx?.Element(c + "title"), c, a);

                    // Legend
                    var leg = chXml.Descendants(c + "legend").FirstOrDefault();
                    chart.LegendPos = leg?.Element(c + "legendPos")?.Attribute("val")?.Value;
                    chart.DataLabels = plotArea?.Descendants(c + "dLbls").Any() == true;

                    // Series
                    foreach (var ser in plotArea?.Descendants().Where(e => e.Name.LocalName == "ser") ?? Enumerable.Empty<XElement>())
                    {
                        var si = new SeriesInfo();
                        // name (tx)
                        var tx = ser.Element(c + "tx");
                        (si.Name, si.NameRef) = ReadChartText(tx, c, a);

                        // categories (cat)
                        var cat = ser.Element(c + "cat");
                        si.CatRef = cat?.Element(c + "strRef")?.Element(c + "f")?.Value
                                    ?? cat?.Element(c + "numRef")?.Element(c + "f")?.Value;

                        // values (val)
                        var val = ser.Element(c + "val");
                        si.ValRef = val?.Element(c + "numRef")?.Element(c + "f")?.Value
                                    ?? val?.Element(c + "strRef")?.Element(c + "f")?.Value;

                        chart.Series.Add(si);
                    }

                    if (!result.TryGetValue(sheetName, out var list)) result[sheetName] = list = new();
                    list.Add(chart);
                }
            }
        }

        return result;

        // helpers
        static (string? txt, string? cellRef) ReadChartText(XElement? node, XNamespace cns, XNamespace ans)
        {
            if (node == null) return (null, null);

            // Accept either <c:title>/<c:tx>… or being passed <c:tx> directly
            var tx = node.Name.LocalName == "tx" ? node : node.Element(cns + "tx");
            if (tx == null) return (null, null);

            var rich = tx.Element(cns + "rich");
            if (rich != null)
            {
                var text = string.Concat(rich.Descendants(ans + "t").Select(t => t.Value));
                return (text, null);
            }

            var strRef = tx.Element(cns + "strRef");
            var f = strRef?.Element(cns + "f")?.Value;
            return (null, f);
        }


        static string DetectChartType(XElement? plotArea)
        {
            if (plotArea == null) return "";
            XNamespace cc = "http://schemas.openxmlformats.org/drawingml/2006/chart";

            XElement? Find(string name) =>
                plotArea.Element(cc + name) ?? plotArea.Element(plotArea.GetDefaultNamespace() + name);

            // Column/Bar share barChart + barDir
            if (Find("barChart") is XElement bc)
            {
                var dirEl = bc.Element(cc + "barDir") ?? bc.Element(plotArea.GetDefaultNamespace() + "barDir");
                var dir = (string?)dirEl?.Attribute("val");
                return string.Equals(dir, "col", StringComparison.OrdinalIgnoreCase) ? "column" : "bar";
            }

            if (Find("lineChart") != null) return "line";
            if (Find("pieChart") != null) return "pie";
            if (Find("pie3DChart") != null) return "pie3D";     // 3-D pie
            if (Find("ofPieChart") != null) return "pie";        // “Pie of Pie” → count as pie
            if (Find("scatterChart") != null) return "scatter";
            if (Find("areaChart") != null) return "area";
            if (Find("doughnutChart") != null) return "doughnut";
            if (Find("radarChart") != null) return "radar";
            if (Find("bubbleChart") != null) return "bubble";
            return "";
        }

    }

    // =====================
    // PIVOT TABLE
    // =====================

    private static CheckResult GradePivotLayout(Rule rule, XLWorkbook wbS)
    {
        var pts = rule.Points;
        var spec = rule.Pivot;
        if (spec is null)
            return new CheckResult("pivot_layout", pts, 0, false, "No pivot spec provided");

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

        // normalize agg tokens
        static string NormAgg(string raw)
        {
            var a = (raw ?? "").ToLowerInvariant();
            if (a.Contains("sum")) return "sum";
            if (a.Contains("count")) return "count";
            if (a.Contains("avg") || a.Contains("average")) return "average";
            if (a.Contains("min")) return "min";
            if (a.Contains("max")) return "max";
            return string.IsNullOrWhiteSpace(a) ? "sum" : a; // default
        }

        var findings = new List<string>();

        foreach (var ws in sheets)
        {
            // ws.PivotTables (unknown concrete type → reflection)
            var pivotsObj = ws.GetType().GetProperty("PivotTables")?.GetValue(ws);
            var pivots = AsEnumerable(pivotsObj);
            if (!pivots.Any()) continue;

            foreach (var pt in pivots)
            {
                var ptName = GetStrProp(pt, "Name") ?? "";

                // optional name filter
                if (!string.IsNullOrWhiteSpace(spec.TableNameLike) &&
                    ptName.IndexOf(spec.TableNameLike, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                // Collect fields using reflection (works across ClosedXML versions)
                HashSet<string> actualRows = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "RowLabels"))
                {
                    actualRows.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));
                }

                HashSet<string> actualCols = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "ColumnLabels"))
                {
                    actualCols.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));
                }

                HashSet<string> actualFilters = new(StringComparer.OrdinalIgnoreCase);
                foreach (var f in GetEnumProp(pt, "ReportFilters"))
                {
                    actualFilters.Add(FirstNonEmpty(GetStrProp(f, "SourceName"), GetStrProp(f, "CustomName"), GetStrProp(f, "Name")));
                }

                HashSet<string> actualValues = new(StringComparer.OrdinalIgnoreCase);
                foreach (var v in GetEnumProp(pt, "Values"))
                {
                    var fieldName = FirstNonEmpty(GetStrProp(v, "SourceName"), GetStrProp(v, "CustomName"), GetStrProp(v, "Name"));
                    if (string.IsNullOrWhiteSpace(fieldName)) continue;

                    // Try SummaryFormula (enum), then Function, then just ToString
                    var sf = GetStrProp(v, "SummaryFormula")
                             ?? GetStrProp(v, "Function")
                             ?? S(v);

                    var agg = NormAgg(sf ?? "");
                    actualValues.Add($"{fieldName}|{agg}");
                }

                // compare
                var missing = new List<string>();
                if (spec.Rows is { Count: > 0 })
                    foreach (var need in spec.Rows) if (!actualRows.Contains(need)) missing.Add($"row '{need}'");
                if (spec.Columns is { Count: > 0 })
                    foreach (var need in spec.Columns) if (!actualCols.Contains(need)) missing.Add($"column '{need}'");
                if (spec.Filters is { Count: > 0 })
                    foreach (var need in spec.Filters) if (!actualFilters.Contains(need)) missing.Add($"filter '{need}'");
                if (spec.Values is { Count: > 0 })
                {
                    foreach (var need in spec.Values)
                    {
                        var wantAgg = NormAgg(need.Agg ?? "sum");
                        if (!actualValues.Contains($"{need.Field}|{wantAgg}"))
                            missing.Add($"value '{need.Field}' with agg '{wantAgg}'");
                    }
                }

                if (missing.Count == 0)
                    return new CheckResult($"pivot_layout:{ws.Name}", pts, pts, true, $"pivot '{ptName}' OK");

                findings.Add($"pivot '{ptName}' missing: {string.Join(", ", missing)}");
            }
        }

        if (findings.Count == 0)
            return new CheckResult("pivot_layout", pts, 0, false, "No pivot tables found or pivot APIs not exposed in this ClosedXML version");
        return new CheckResult("pivot_layout", pts, 0, false, string.Join(" | ", findings));
    }

    // =====================
    // CONDITIONAL FORMATTING
    // =====================

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

    private static CheckResult GradeConditionalFormat(Rule rule, XLWorkbook wbS)
    {
        var pts = rule.Points;
        var spec = rule.Cond ?? new ConditionalFormatSpec();
        var sheetName = spec.Sheet ?? wbS.Worksheets.First().Name;

        if (_studentZipBytes is null)
            return new CheckResult("conditional_format", pts, 0, false,
                "Student .xlsx bytes not available to inspect conditional formats");

        var expected = DescribeCond(spec);
        var ok = FindCFInStudentZip_NoClosedXml(_studentZipBytes, sheetName, spec,
                                                out var reason, out var matchedSummary);

        var note = ok
            ? $"matched: {matchedSummary}"
            : $"expected: {expected}; {reason}";

        return new CheckResult($"conditional_format:{sheetName}", pts, ok ? pts : 0, ok, note);
    }


    private static bool FindCFInStudentZip_NoClosedXml(
    byte[] zipBytes, string sheetName, ConditionalFormatSpec spec,
    out string reason, out string matchedSummary)
    {
        reason = "no matching conditional format";
        matchedSummary = "";
        try
        {
            using var ms = new MemoryStream(zipBytes);
            using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);

            var wbEntry = zip.GetEntry("xl/workbook.xml");
            if (wbEntry is null) { reason = "workbook.xml missing"; return false; }

            var wbXml = System.Xml.Linq.XDocument.Load(wbEntry.Open());
            System.Xml.Linq.XName S(string n) => System.Xml.Linq.XName.Get(n, "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            // sheet name -> index
            var indexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int idx = 1;
            var sheetsEl = wbXml.Root?.Element(S("sheets"));
            if (sheetsEl == null) { reason = "no <sheets>"; return false; }
            foreach (var sh in sheetsEl.Elements(S("sheet")))
            {
                var nm = (string?)sh.Attribute("name") ?? $"Sheet{idx}";
                indexByName[nm] = idx++;
            }
            if (!indexByName.TryGetValue(sheetName, out var sheetIdx)) { reason = $"sheet '{sheetName}' not found"; return false; }

            var sheetPath = $"xl/worksheets/sheet{sheetIdx}.xml";
            var sheetEntry = zip.GetEntry(sheetPath); if (sheetEntry is null) { reason = $"{sheetPath} missing"; return false; }
            var wsXml = System.Xml.Linq.XDocument.Load(sheetEntry.Open());

            var stylesEntry = zip.GetEntry("xl/styles.xml");
            System.Xml.Linq.XDocument? stylesXml = stylesEntry != null ? System.Xml.Linq.XDocument.Load(stylesEntry.Open()) : null;

            static string Norm(string? s)
            {
                if (string.IsNullOrWhiteSpace(s)) return "";
                s = s.Trim();
                if (s.StartsWith("=")) s = s.Substring(1);
                return s.Replace(" ", "");
            }

            string? ExtractFillRgb(System.Xml.Linq.XDocument? styles, int? dxfId)
            {
                if (styles is null || dxfId is null || dxfId < 0) return null;
                var dxfs = styles.Root?.Element(S("dxfs"))?.Elements(S("dxf")).ToList();
                if (dxfs == null || dxfId.Value >= dxfs.Count) return null;
                var dxf = dxfs[dxfId.Value];
                var rgb = dxf.Element(S("fill"))?.Element(S("patternFill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value
                       ?? dxf.Element(S("fill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value;
                if (!string.IsNullOrWhiteSpace(rgb) && rgb.Length == 8) rgb = rgb.Substring(2); // strip ARGB alpha
                return rgb;
            }

            // Collect for diagnostics on failure
            var seenSummaries = new List<string>();
            int totalRules = 0, overlapRules = 0;

            bool RangeOverlapsA1(string sqref, string? expectedRange)
            {
                if (string.IsNullOrWhiteSpace(expectedRange)) return true;
                if (!TryParseA1Range(expectedRange, out var want)) return false;
                foreach (var token in sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                    if (TryParseA1Range(token, out var got) && RectsIntersect(want, got)) return true;
                return false;
            }

            var cfBlocks = wsXml.Root!.Elements(S("conditionalFormatting")).ToList();
            foreach (var block in cfBlocks)
            {
                var sqref = (string?)block.Attribute("sqref") ?? "";
                foreach (var ruleEl in block.Elements(S("cfRule")))
                {
                    totalRules++;

                    var t = (string?)ruleEl.Attribute("type");
                    var opXml = (string?)ruleEl.Attribute("operator");
                    var opTok = MapXmlOp(opXml);
                    var frms = ruleEl.Elements(S("formula")).Select(e => e.Value?.Trim()).ToList();
                    var txt = (string?)ruleEl.Attribute("text");

                    int? dxfId = int.TryParse((string?)ruleEl.Attribute("dxfId"), out var _id) ? _id : (int?)null;
                    var fillRgb = ExtractFillRgb(stylesXml, dxfId);

                    var summary = DescribeCondFromXml(sheetName, sqref, t, opTok, frms, txt, fillRgb);
                    if (RangeOverlapsA1(sqref, spec.Range)) { overlapRules++; seenSummaries.Add(summary); }

                    // matching logic
                    if (!RangeOverlapsA1(sqref, spec.Range)) continue;
                    if (spec.Type != null && !string.Equals(t ?? "", spec.Type, StringComparison.OrdinalIgnoreCase)) continue;
                    if (spec.Type == "cellIs" && spec.Operator != null &&
                        !string.Equals(opTok ?? "", spec.Operator, StringComparison.OrdinalIgnoreCase)) continue;

                    if (spec.Text != null && (txt == null || !txt.Contains(spec.Text, StringComparison.OrdinalIgnoreCase))) continue;

                    string NF(string? s) => Norm(s);
                    if (spec.Formula1 != null && (frms.Count < 1 || NF(frms[0]) != NF(spec.Formula1))) continue;
                    if (spec.Formula2 != null && (frms.Count < 2 || NF(frms[1]) != NF(spec.Formula2))) continue;

                    if (!string.IsNullOrWhiteSpace(spec.FillRgb) && !string.IsNullOrWhiteSpace(fillRgb))
                    {
                        if (!fillRgb.EndsWith(spec.FillRgb!, StringComparison.OrdinalIgnoreCase)) continue;
                    }

                    matchedSummary = summary;
                    reason = "matched";
                    return true;
                }
            }

            // Build descriptive failure note
            if (totalRules == 0)
            {
                reason = "no conditional formatting rules on this sheet";
            }
            else if (overlapRules == 0)
            {
                reason = $"no rules applied to expected range {spec.Range}";
            }
            else
            {
                var show = string.Join(" | ", seenSummaries.Take(3));
                reason = $"no rule matched exactly; {overlapRules} rule(s) apply to the range; closest examples: {show}";
            }

            return false;
        }
        catch (Exception ex)
        {
            reason = ex.Message;
            matchedSummary = "";
            return false;
        }
    }

    // =====================
    // EXCEL TABLE
    // =====================

    private static CheckResult GradeTable(Rule rule, IXLWorksheet wsS)
    {
        var pts = rule.Points;
        var spec = rule.Table;
        if (spec is null)
            return new CheckResult("table", pts, 0, false, "No table spec provided");

        // Sheet gating
        if (!string.IsNullOrWhiteSpace(spec.Sheet) &&
            !string.Equals(spec.Sheet, wsS.Name, StringComparison.OrdinalIgnoreCase))
        {
            return new CheckResult("table", pts, 0, false, $"Expected on sheet '{spec.Sheet}', grading '{wsS.Name}'");
        }

        // Candidate tables
        var tables = wsS.Tables.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(spec.NameLike))
            tables = tables.Where(t => t.Name.IndexOf(spec.NameLike!, StringComparison.OrdinalIgnoreCase) >= 0);

        if (!tables.Any())
            return new CheckResult("table", pts, 0, false,
                $"No table {(string.IsNullOrWhiteSpace(spec.NameLike) ? "" : $"matching '{spec.NameLike}' ")}found on '{wsS.Name}'");

        int checksTotal = 0, bestHits = -1;
        string bestNote = "", bestName = "";

        foreach (var t in tables)
        {
            var notes = new List<string>();
            int cks = 0, hits = 0;

            // ---------- HEADERS ----------
            var headers = t.Fields.Select(f => f.Name?.Trim() ?? "").ToList();
            if (spec.Columns is { Count: > 0 })
            {
                foreach (var want in spec.Columns)
                {
                    cks++;
                    bool present = headers.Any(h => string.Equals(h, want, StringComparison.OrdinalIgnoreCase));
                    if (present) { hits++; notes.Add($"[{t.Name}] has '{want}'"); }
                    else notes.Add($"[{t.Name}] missing '{want}'");
                }
                if (spec.RequireOrder == true)
                {
                    cks++;
                    bool orderOk = true; int last = -1;
                    foreach (var want in spec.Columns)
                    {
                        int idx = headers.FindIndex(h => string.Equals(h, want, StringComparison.OrdinalIgnoreCase));
                        if (idx < 0 || idx < last) { orderOk = false; break; }
                        last = idx;
                    }
                    if (orderOk) { hits++; notes.Add($"[{t.Name}] column order ok"); }
                    else notes.Add($"[{t.Name}] column order wrong");
                }
            }

            // ---------- RANGE REF (full table range, incl. header) ----------
            if (!string.IsNullOrWhiteSpace(spec.RangeRef))
            {
                cks++;
                bool ok = false;
                try
                {
                    string sheetPart = wsS.Name, addrPart = spec.RangeRef!;
                    int bang = spec.RangeRef!.IndexOf('!');
                    if (bang >= 0) { sheetPart = spec.RangeRef!.Substring(0, bang); addrPart = spec.RangeRef!.Substring(bang + 1); }
                    var sh = string.Equals(sheetPart, wsS.Name, StringComparison.OrdinalIgnoreCase)
                        ? wsS
                        : wsS.Workbook.Worksheet(sheetPart);

                    var expected = sh.Range(addrPart).RangeAddress;
                    var got = t.RangeAddress;
                    ok = got.FirstAddress.RowNumber == expected.FirstAddress.RowNumber
                         && got.FirstAddress.ColumnNumber == expected.FirstAddress.ColumnNumber
                         && got.LastAddress.RowNumber == expected.LastAddress.RowNumber
                         && got.LastAddress.ColumnNumber == expected.LastAddress.ColumnNumber;
                }
                catch { ok = false; }

                if (ok) { hits++; notes.Add($"[{t.Name}] range matches {spec.RangeRef}"); }
                else notes.Add($"[{t.Name}] range != {spec.RangeRef} (got {t.RangeAddress.ToStringRelative()})");
            }

            // ---------- DIMENSIONS (data body only) ----------
            var body = t.DataRange;
            int bodyRows = body?.RowCount() ?? 0;
            int bodyCols = body?.ColumnCount() ?? 0;

            if (spec.Rows.HasValue)
            {
                cks++;
                bool ok = (spec.AllowExtraRows == true) ? (bodyRows >= spec.Rows.Value) : (bodyRows == spec.Rows.Value);
                if (ok) { hits++; notes.Add($"rows {bodyRows} ok"); }
                else notes.Add($"rows {bodyRows} not {(spec.AllowExtraRows == true ? ">=" : "=")} {spec.Rows}");
            }
            if (spec.Cols.HasValue)
            {
                cks++;
                bool ok = (spec.AllowExtraCols == true) ? (bodyCols >= spec.Cols.Value) : (bodyCols == spec.Cols.Value);
                if (ok) { hits++; notes.Add($"cols {bodyCols} ok"); }
                else notes.Add($"cols {bodyCols} not {(spec.AllowExtraCols == true ? ">=" : "=")} {spec.Cols}");
            }

            // ---------- CONTAINS ROWS (subset match) ----------
            if (spec.ContainsRows is { Count: > 0 })
            {
                var idxByName = headers
                    .Select((h, i) => (h, i))
                    .ToDictionary(x => x.h, x => x.i, StringComparer.OrdinalIgnoreCase);

                foreach (var required in spec.ContainsRows)
                {
                    cks++;
                    bool found = false;
                    if (body != null)
                    {
                        foreach (var row in body.Rows())
                        {
                            bool match = true;
                            foreach (var kv in required)
                            {
                                if (!idxByName.TryGetValue(kv.Key, out int ci)) { match = false; break; }
                                var text = row.Cell(ci + 1).GetFormattedString()?.Trim() ?? "";
                                if (!string.Equals(text, (kv.Value ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                                { match = false; break; }
                            }
                            if (match) { found = true; break; }
                        }
                    }
                    if (found) { hits++; notes.Add($"contains: {string.Join(", ", required.Select(kv => $"{kv.Key}='{kv.Value}'"))}"); }
                    else notes.Add($"missing: {string.Join(", ", required.Select(kv => $"{kv.Key}='{kv.Value}'"))}");
                }
            }

            // ---------- WHOLE-BODY COMPARISON ----------
            if (spec.BodyMatch == true && spec.BodyRows is { Count: > 0 })
            {
                cks++;

                // student values (formatted)
                var sBody = new List<List<string>>();
                if (body != null)
                {
                    foreach (var r in body.Rows())
                    {
                        var rowVals = new List<string>();
                        foreach (var c in r.Cells())
                            rowVals.Add(c.GetFormattedString() ?? "");
                        sBody.Add(rowVals);
                    }
                }

                bool trim = spec.BodyTrim != false;           // default true
                bool caseSens = spec.BodyCaseSensitive == true;
                string Norm(string x) => trim ? (x ?? "").Trim() : (x ?? "");
                StringComparer cmp = caseSens ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

                // shape must match
                bool shapeOk = sBody.Count == spec.BodyRows.Count &&
                               (sBody.Count == 0 || sBody[0].Count == spec.BodyRows[0].Count);

                bool match = shapeOk;
                if (match)
                {
                    if (spec.BodyOrderMatters == true)
                    {
                        for (int i = 0; i < sBody.Count && match; i++)
                            for (int j = 0; j < sBody[i].Count && match; j++)
                                if (!cmp.Equals(Norm(sBody[i][j]), Norm(spec.BodyRows[i][j])))
                                    match = false;
                    }
                    else
                    {
                        string Key(List<string> row) => string.Join("\u001F", row.Select(Norm));
                        var left = sBody.Select(Key).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
                        var right = spec.BodyRows.Select(Key).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
                        match = left.Count == right.Count && left.All(kv => right.TryGetValue(kv.Key, out int n) && n == kv.Value);
                    }
                }

                if (match) { hits++; notes.Add("body matches"); }
                else notes.Add("body does not match");
            }

            // Best candidate scoring
            if (cks > 0 && hits > bestHits)
            {
                bestHits = hits;
                checksTotal = cks;
                bestName = t.Name;
                bestNote = string.Join(" | ", notes);
            }
        }

        if (checksTotal == 0)
            return new CheckResult("table", pts, 0, false, "No checks declared (add columns / range_ref / rows/cols / contains_rows / body_match).");

        double frac = (double)bestHits / checksTotal;
        double earned = pts * frac;
        bool pass = Math.Abs(frac - 1.0) < 1e-9;
        return new CheckResult($"table:{bestName}", pts, earned, pass, bestNote);
    }



    // ---------------- A1 parsing helpers (no ClosedXML needed) ------------------

    private static bool RangeOverlapsA1(string sqref, string? expectedRange)
    {
        if (string.IsNullOrWhiteSpace(expectedRange)) return true; // no constraint
        if (!TryParseA1Range(expectedRange, out var R1)) return false;

        foreach (var token in sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (TryParseA1Range(token, out var R2))
                if (RectsIntersect(R1, R2)) return true;
        }
        return false;
    }

    private static bool RectsIntersect((int r1, int c1, int r2, int c2) a, (int r1, int c1, int r2, int c2) b)
    {
        return !(b.c1 > a.c2 || b.c2 < a.c1 || b.r1 > a.r2 || b.r2 < a.r1);
    }

    private static bool TryParseA1Range(string a1, out (int r1, int c1, int r2, int c2) rect)
    {
        // returns 1-based inclusive coordinates
        rect = (0, 0, 0, 0);
        if (string.IsNullOrWhiteSpace(a1)) return false;

        string s = a1.Replace("$", "").Trim();
        // Whole column: "B:B"
        var parts = s.Split(':');
        if (parts.Length == 2 && IsLetters(parts[0]) && IsLetters(parts[1]))
        {
            int c1 = ColToNum(parts[0]); int c2 = ColToNum(parts[1]);
            if (c1 > c2) (c1, c2) = (c2, c1);
            rect = (1, c1, 1_048_576, c2); // Excel max rows
            return true;
        }
        // Whole row: "2:2"
        if (parts.Length == 2 && IsDigits(parts[0]) && IsDigits(parts[1]))
        {
            int r1 = int.Parse(parts[0]); int r2 = int.Parse(parts[1]);
            if (r1 > r2) (r1, r2) = (r2, r1);
            rect = (r1, 1, r2, 16_384); // Excel max columns
            return true;
        }
        // Cell range: "A1:B10"
        if (parts.Length == 2 && TryParseCell(parts[0], out var A) && TryParseCell(parts[1], out var B))
        {
            int r1 = Math.Min(A.r, B.r), r2 = Math.Max(A.r, B.r);
            int c1 = Math.Min(A.c, B.c), c2 = Math.Max(A.c, B.c);
            rect = (r1, c1, r2, c2);
            return true;
        }
        // Single cell: "C7"
        if (parts.Length == 1 && TryParseCell(parts[0], out var C))
        {
            rect = (C.r, C.c, C.r, C.c);
            return true;
        }

        return false;

        static bool IsLetters(string x) => x.All(ch => ch is >= 'A' and <= 'Z' || ch is >= 'a' and <= 'z');
        static bool IsDigits(string x) => x.All(ch => char.IsDigit(ch));

        static bool TryParseCell(string s, out (int r, int c) cell)
        {
            cell = (0, 0);
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.ToUpperInvariant();

            int i = 0;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            if (i == 0 || i == s.Length) return false;

            string col = s.Substring(0, i);
            string row = s.Substring(i);
            if (!int.TryParse(row, out var r)) return false;

            cell = (r, ColToNum(col));
            return true;
        }

        static int ColToNum(string col)
        {
            col = col.ToUpperInvariant();
            int n = 0;
            foreach (var ch in col)
            {
                if (ch < 'A' || ch > 'Z') continue;
                n = n * 26 + (ch - 'A' + 1);
            }
            return n;
        }
    }



    // =====================
    // CUSTOM
    // =====================

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

    // =====================
    // HELPERS
    // =====================

    private static (bool ok, string reason) AnyOfMatch(List<RuleOption> options, Func<RuleOption, (bool ok, string reason)> check)
    {
        var reasons = new List<string>();
        foreach (var opt in options)
        {
            var (ok, reason) = check(opt);
            if (ok) return (true, "Matched one acceptable option.");
            reasons.Add(reason);
        }
        return (false, "No acceptable option matched. " + string.Join(" | ", reasons));
    }

    private static string Normalize(object? o) => (o?.ToString() ?? "").Trim();

    private static bool TryToDouble(object? o, out double d)
    {
        if (o is null) { d = 0; return false; }
        switch (o)
        {
            case double dx: d = dx; return true;
            case float f: d = f; return true;
            case int i: d = i; return true;
            case long l: d = l; return true;
            case decimal m: d = (double)m; return true;
            case DateTime dt: d = dt.ToOADate(); return true;
            default:
                var s = o.ToString();
                if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                { d = v; return true; }
                if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out v))
                { d = v; return true; }
                d = 0; return false;
        }
    }

    private static object? JsonToNet(JsonElement e) => e.ValueKind switch
    {
        JsonValueKind.String => e.GetString(),
        JsonValueKind.Number => e.TryGetInt64(out var i) ? i : e.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        _ => e.ToString()
    };

    // Helper (unchanged): normalize for robust comparison
    private static string NormalizeFormula(string? f)
    {
        var s = (f ?? "").Trim();
        if (s.Length == 0) return "";
        if (s[0] != '=') s = "=" + s;
        s = s.Replace(" ", "").Replace("$", "");
        return s.ToUpperInvariant();
    }


    private static string XLColorToArgb(XLColor color)
    {
        // ClosedXML 0.105.x: XLColor.Color is System.Drawing.Color
        var sys = color.Color;
        return $"FF{sys.R:X2}{sys.G:X2}{sys.B:X2}";
    }

    private static string NormalizeArgb(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var up = s.Trim().ToUpperInvariant();
        if (up.Length == 6) return "FF" + up;          // RGB → ARGB
        if (up.Length == 8) return up;                 // already ARGB
        if (up.Length == 9 && up.StartsWith("#")) return up[1..];
        if (up.Length == 7 && up.StartsWith("#")) return "FF" + up[1..];
        return up;
    }

    // Human labels for operator tokens
    private static string OpLabel(string? op) => op switch
    {
        "gt" => "> greater than",
        "ge" => "≥ greater than or equal",
        "lt" => "< less than",
        "le" => "≤ less than or equal",
        "eq" => "= equal to",
        "ne" => "≠ not equal to",
        "between" => "between (inclusive)",
        "notBetween" => "not between",
        _ => op ?? ""
    };

    private static string DescribeCond(ConditionalFormatSpec s)
    {
        var bits = new List<string>();
        if (!string.IsNullOrWhiteSpace(s.Type))
            bits.Add(s.Type == "cellIs" ? "Cell is…" :
                     s.Type == "expression" ? "Formula (TRUE/FALSE)" :
                     s.Type == "containsText" ? "Contains text" : s.Type!);

        if (!string.IsNullOrWhiteSpace(s.Operator) && s.Type == "cellIs")
            bits.Add(OpLabel(s.Operator));

        if (!string.IsNullOrWhiteSpace(s.Text))
            bits.Add($"text \"{s.Text}\"");

        if (!string.IsNullOrWhiteSpace(s.Formula1))
            bits.Add($"F1: {s.Formula1}");

        if (!string.IsNullOrWhiteSpace(s.Formula2))
            bits.Add($"F2: {s.Formula2}");

        if (!string.IsNullOrWhiteSpace(s.FillRgb))
            bits.Add($"fill #{s.FillRgb}");

        if (!string.IsNullOrWhiteSpace(s.Range))
            bits.Add($"range {s.Range}");

        return string.Join(", ", bits);
    }

    private static string DescribeCondFromXml(string sheet, string sqref, string? type, string? opToken, IList<string> frms, string? text, string? fillRgb)
    {
        var s = new ConditionalFormatSpec
        {
            Sheet = sheet,
            Range = sqref.Split(' ').FirstOrDefault(),
            Type = type,
            Operator = opToken,
            Formula1 = frms.ElementAtOrDefault(0),
            Formula2 = frms.ElementAtOrDefault(1),
            Text = text,
            FillRgb = fillRgb
        };
        return DescribeCond(s);
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



    private static bool ArgbEqual(string a, string b) => NormalizeArgb(a) == NormalizeArgb(b);
}
