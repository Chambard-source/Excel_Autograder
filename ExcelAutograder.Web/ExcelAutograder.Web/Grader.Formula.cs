using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Grades a formula in a student worksheet cell against allowed expectations.
    /// Supports:
    /// <list type="bullet">
    ///   <item><description>Literal match (<see cref="Rule.ExpectedFormula"/>)</description></item>
    ///   <item><description>Regex match (<see cref="Rule.ExpectedFormulaRegex"/>)</description></item>
    ///   <item><description>Match the key workbook’s formula (<see cref="Rule.ExpectedFromKey"/>)</description></item>
    ///   <item><description>Multiple options via <see cref="Rule.AnyOf"/></description></item>
    ///   <item><description>Required absolute refs (<see cref="Rule.RequireAbsolute"/>)</description></item>
    ///   <item><description>Value check via <c>ValueMatches</c> (must match expected value too)</description></item>
    /// </list>
    /// </summary>
    /// <param name="rule">
    /// The rubric rule defining the target cell, the expectations (single or <c>AnyOf</c>),
    /// required absolutes, and points.
    /// </param>
    /// <param name="wsS">Student worksheet.</param>
    /// <param name="wsK">Key worksheet (optional; required if using <c>ExpectedFromKey</c>).</param>
    /// <returns>
    /// A <see cref="CheckResult"/> awarding:
    /// <list type="bullet">
    ///   <item><description>Full credit if formula content matches and value matches.</description></item>
    ///   <item><description>Half credit if value matches but formula content does not.</description></item>
    ///   <item><description>Half credit if content matches but required absolutes are missing.</description></item>
    ///   <item><description>Zero otherwise (with reasons).</description></item>
    /// </returns>
    /// <remarks>
    /// The student’s raw formula (A1 then R1C1 fallback) is normalized via <c>NormalizeFormula</c>
    /// before content comparisons. Expected absolutes are validated either against the expected
    /// literal/key formula or via <see cref="MissingAbsoluteRefs(string)"/> for regex cases.
    /// </remarks>
    private static CheckResult GradeFormula(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var cellAddr = rule.Cell ?? throw new Exception("formula check missing 'cell'");
        var pts = rule.Points;

        // Local helper: tries to parse a number possibly formatted with % or commas.
        static bool TryParseNumber(string s, out double v)
        {
            var raw = s.Trim();
            var isPct = raw.EndsWith("%", StringComparison.Ordinal);
            raw = raw.Replace("%", "").Replace(",", "");
            raw = new string(raw.Where(ch => char.IsDigit(ch) || ch == '-' || ch == '.').ToArray());

            var ok = double.TryParse(raw, System.Globalization.NumberStyles.Float,
                                     System.Globalization.CultureInfo.InvariantCulture, out v);
            if (ok && isPct) v /= 100.0;
            return ok;
        }

        var sCell = wsS.Cell(cellAddr);
        var a1 = sCell.FormulaA1;
        var r1c1 = sCell.FormulaR1C1;

        // If no formula at all, allow a value-only partial check, but no credit if formula is required.
        var hasRealFormula = sCell.HasFormula || (!string.IsNullOrWhiteSpace(a1) || !string.IsNullOrWhiteSpace(r1c1));
        if (!hasRealFormula)
        {
            var (valOk0, valDetail0) = ValueMatches(sCell, rule);
            var msg0 = valOk0
                ? $"cell has no formula (value matches{(string.IsNullOrWhiteSpace(valDetail0) ? "" : $": {valDetail0}")}; no credit)"
                : "cell has no formula";
            return new CheckResult($"formula:{cellAddr}", pts, 0, false, msg0);
        }

        string sRaw = a1 ?? r1c1 ?? "";
        string sNorm = NormalizeFormula(sRaw);

        var opts = (rule.AnyOf is { Count: > 0 }
            ? rule.AnyOf.Select(o => (o.ExpectedFormula, o.ExpectedFormulaRegex, o.ExpectedFromKey, o.RequireAbsolute, "option"))
            : new[] { (rule.ExpectedFormula, rule.ExpectedFormulaRegex, rule.ExpectedFromKey, rule.RequireAbsolute, "rule") }
        ).ToList();

        var reasons = new List<string>();
        string? expectedHint = null;

        foreach (var (litExp, rxExp, fromKey, requireAbs, origin) in opts)
        {
            // 1) Literal expected
            var litNorm = NormalizeFormula(litExp ?? "");
            if (!string.IsNullOrEmpty(litNorm))
            {
                if (string.IsNullOrWhiteSpace(expectedHint)) expectedHint = litExp;
                bool contentOk = (sNorm == litNorm);

                if (!contentOk)
                {
                    reasons.Add($"formula mismatch | expected='{litExp}' ({origin}); got='{sRaw}'");
                    continue;
                }

                if (requireAbs == true)
                {
                    var missing = MissingAbsolutesFromExpected(litExp!, sRaw);
                    if (missing.Count > 0)
                    {
                        var partial = Math.Round(pts * 0.5, 3);
                        return new CheckResult(
                            $"formula:{cellAddr}", pts, partial, false,
                            $"formula correct; missing required absolutes: {string.Join(", ", missing)} | got='{sRaw}'"
                        );
                    }
                }

                var (ok, detail) = ValueMatches(sCell, rule);
                if (!ok)
                {
                    return new CheckResult(
                        $"formula:{cellAddr}", pts, 0, false,
                        $"formula correct but value mismatch: {detail} | got formula '{sRaw}'"
                    );
                }

                return new CheckResult(
                    $"formula:{cellAddr}", pts, pts, true,
                    "formula correct" + (requireAbs == true ? " with required absolutes" : "") +
                    (string.IsNullOrWhiteSpace(detail) ? "" : $" | {detail}")
                );
            }

            // 2) Regex expected
            if (!string.IsNullOrWhiteSpace(rxExp))
            {
                if (string.IsNullOrWhiteSpace(expectedHint)) expectedHint = $"/{rxExp}/";
                bool contentOk = Regex.IsMatch(sNorm, $"^{rxExp}$", RegexOptions.IgnoreCase);

                if (!contentOk)
                {
                    reasons.Add($"formula mismatch (regex) | regex='{rxExp}' ({origin}); got='{sRaw}'");
                    continue;
                }

                if (requireAbs == true)
                {
                    var absInfo = MissingAbsoluteRefs(sRaw);
                    if (!absInfo.ok)
                    {
                        var partial = Math.Round(pts * 0.5, 3);
                        return new CheckResult(
                            $"formula:{cellAddr}", pts, partial, false,
                            $"formula correct; missing required absolutes: {string.Join(", ", absInfo.missing)} | got='{sRaw}'"
                        );
                    }
                }

                var (ok, detail) = ValueMatches(sCell, rule);
                if (!ok)
                {
                    return new CheckResult(
                        $"formula:{cellAddr}", pts, 0, false,
                        $"formula correct (regex) but value mismatch: {detail} | got formula '{sRaw}'"
                    );
                }

                return new CheckResult(
                    $"formula:{cellAddr}", pts, pts, true,
                    "formula correct (regex)" + (requireAbs == true ? " with required absolutes" : "") +
                    (string.IsNullOrWhiteSpace(detail) ? "" : $" | {detail}")
                );
            }

            // 3) Expected from key
            if (fromKey == true)
            {
                if (wsK is null)
                {
                    reasons.Add($"key sheet missing ({origin})");
                    continue;
                }

                var kc = wsK.Cell(cellAddr);
                var kRaw = kc.HasFormula ? (kc.FormulaA1 ?? kc.FormulaR1C1 ?? "") : "";
                var kNorm = NormalizeFormula(kRaw);

                if (string.IsNullOrWhiteSpace(expectedHint) && !string.IsNullOrWhiteSpace(kRaw))
                    expectedHint = kRaw;

                if (string.IsNullOrEmpty(kNorm))
                {
                    reasons.Add("key cell has no formula");
                    continue;
                }

                bool contentOk = (sNorm == kNorm);

                if (!contentOk)
                {
                    reasons.Add($"formula mismatch (from key) | expected='{kRaw}' ; got='{sRaw}'");
                    continue;
                }

                if (requireAbs == true)
                {
                    var missing = MissingAbsolutesFromExpected(kRaw, sRaw);
                    if (missing.Count > 0)
                    {
                        var partial = Math.Round(pts * 0.5, 3);
                        return new CheckResult(
                            $"formula:{cellAddr}", pts, partial, false,
                            $"formula correct; missing required absolutes: {string.Join(", ", missing)} | got='{sRaw}'"
                        );
                    }
                }

                var (ok, detail) = ValueMatches(sCell, rule);
                if (!ok)
                {
                    return new CheckResult(
                        $"formula:{cellAddr}", pts, 0, false,
                        $"formula correct (matches key) but value mismatch: {detail} | got formula '{sRaw}'"
                    );
                }

                return new CheckResult(
                    $"formula:{cellAddr}", pts, pts, true,
                    "formula correct (matches key)" + (requireAbs == true ? " with required absolutes" : "") +
                    (string.IsNullOrWhiteSpace(detail) ? "" : $" | {detail}")
                );
            }

            reasons.Add($"no expected provided ({origin})");
        }

        // If we got here, content didn't match any path. Grant partial only if the value is correct.
        {
            var (valOk, valDetail) = ValueMatches(sCell, rule);
            if (valOk)
            {
                var partial = Math.Round(pts * 0.5, 3);
                var expectedBit = string.IsNullOrWhiteSpace(expectedHint) ? "" : $" | expected='{expectedHint}'";
                return new CheckResult(
                    $"formula:{cellAddr}", pts, partial, false,
                    $"value correct but formula incorrect{(string.IsNullOrWhiteSpace(valDetail) ? "" : ": " + valDetail)} | got formula '{sRaw}'{expectedBit}"
                );
            }
        }

        var failMsg = string.Join(" | ", reasons);
        return new CheckResult($"formula:{cellAddr}", pts, 0, false, failMsg);
    }

    // ===== helpers kept with formula for locality =====

    /// <summary>
    /// Detects whether any absolute column or row references (e.g., <c>$A$1</c>) are present in a formula.
    /// </summary>
    /// <param name="formulaA1">A1 formula text.</param>
    /// <returns>
    /// <c>(any, col, row)</c> indicating presence of any absolute, absolute column(s), and absolute row(s).
    /// </returns>
    private static (bool any, bool col, bool row) InspectAbsoluteRefs(string? formulaA1)
    {
        var f = formulaA1 ?? "";
        bool col = Regex.IsMatch(f, @"\$[A-Za-z]");
        bool row = Regex.IsMatch(f, @"\$\d");
        return (col || row, col, row);
    }

    /// <summary>
    /// Finds cell references in an A1 formula that are not fully absolute (i.e., not <c>$Col$Row</c>).
    /// </summary>
    /// <param name="formulaA1">A1 formula text.</param>
    /// <returns>
    /// <c>(ok, missing)</c> where <c>ok</c> is true if all references are fully absolute,
    /// and <c>missing</c> lists the tokens that are not (sheet prefixes removed).
    /// </returns>
    private static (bool ok, List<string> missing) MissingAbsoluteRefs(string? formulaA1)
    {
        var text = formulaA1 ?? string.Empty;
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
                if (bang >= 0 && bang + 1 < token.Length) token = token[(bang + 1)..];
                missing.Add(token);
            }
        }
        return (missing.Count == 0, missing);
    }

    /// <summary>
    /// Extracts cell reference endpoints (with whether column/row are absolute) from an A1 formula.
    /// </summary>
    /// <param name="formulaA1">A1 formula text.</param>
    /// <returns>
    /// List of tuples: <c>(colAbs, rowAbs, token)</c> where <c>token</c> is the reference without any sheet prefix.
    /// </returns>
    private static List<(bool colAbs, bool rowAbs, string token)> ExtractEndpoints(string? formulaA1)
    {
        var res = new List<(bool colAbs, bool rowAbs, string token)>();
        if (string.IsNullOrWhiteSpace(formulaA1)) return res;

        var rx = new Regex(@"(?<![A-Z0-9_])(?:'[^']+'!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)", RegexOptions.IgnoreCase);
        foreach (Match m in rx.Matches(formulaA1))
        {
            bool colAbs = m.Groups[1].Value == "$";
            bool rowAbs = m.Groups[3].Value == "$";
            string token = m.Value;
            int bang = token.LastIndexOf('!');
            if (bang >= 0 && bang + 1 < token.Length) token = token[(bang + 1)..];
            res.Add((colAbs, rowAbs, token));
        }
        return res;
    }

    /// <summary>
    /// Compares the expected and student formulas reference-by-reference and returns the tokens
    /// where the student is missing absolutes that exist in the expected formula.
    /// </summary>
    /// <param name="expectedA1">Expected A1 formula (from literal or key).</param>
    /// <param name="studentA1">Student A1 formula.</param>
    /// <returns>Distinct list of tokens missing required absolutes.</returns>
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

        for (int i = got.Count; i < exp.Count; i++)
            if (exp[i].colAbs || exp[i].rowAbs) missing.Add(exp[i].token);

        return missing.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }
}
