using System;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Grades cell values over a range against either a key workbook's values (same addresses)
    /// or a single literal expected value. Numeric values are compared with absolute tolerance;
    /// otherwise trimmed string equality is used.
    /// </summary>
    /// <param name="rule">
    /// Requires <c>Range</c>. Honors <c>Points</c>, <c>Tolerance</c>, <c>ExpectedFromKey</c>, and <c>Expected</c>.
    /// </param>
    /// <param name="wsS">Student worksheet.</param>
    /// <param name="wsK">Key worksheet (used when <c>ExpectedFromKey</c> is true).</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>range_value:{range}</c>, awarding proportional credit
    /// equal to matchedCells / totalCells.
    /// </returns>
    /// <exception cref="Exception">Thrown when <c>rule.Range</c> is missing.</exception>
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

    /// <summary>
    /// Grades formulas over a range. Compares against:
    /// (1) the key workbook's formulas at the same addresses if <c>ExpectedFromKey</c> is true;
    /// (2) a regex pattern; or (3) a literal expected formula.
    /// Normalization enforces leading '=', strips spaces and '$', and upper-cases.
    /// </summary>
    /// <param name="rule">
    /// Requires <c>Range</c>. Honors <c>ExpectedFromKey</c>, <c>ExpectedFormulaRegex</c>, and <c>ExpectedFormula</c>.
    /// </param>
    /// <param name="wsS">Student worksheet.</param>
    /// <param name="wsK">Key worksheet (used when comparing to key formulas).</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>range_formula:{range}</c>, awarding proportional credit
    /// equal to matchedFormulas / totalCells.
    /// </returns>
    /// <exception cref="Exception">Thrown when <c>rule.Range</c> is missing.</exception>
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

    /// <summary>
    /// Grades formatting over a range by evaluating each used cell against one or more
    /// acceptable <see cref="FormatSpec"/> options (<c>AnyOf</c> support). Tally is proportional.
    /// </summary>
    /// <param name="rule">Requires <c>Range</c>. Honors <c>Format</c> and <c>AnyOf[].Format</c>.</param>
    /// <param name="wsS">Student worksheet.</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>range_format:{range}</c>, awarding proportional credit
    /// equal to formattedMatches / totalUsedCells.
    /// </returns>
    /// <exception cref="Exception">Thrown when <c>rule.Range</c> is missing.</exception>
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

    /// <summary>
    /// Verifies that a range contains a numeric sequence defined by <c>Start</c> and <c>Step</c>.
    /// Blank cells are counted and must match the expected numeric for full credit.
    /// </summary>
    /// <param name="rule">Requires <c>Range</c>. Honors <c>Start</c> (default 1.0) and <c>Step</c> (default 1.0).</param>
    /// <param name="wsS">Student worksheet.</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>range_sequence:{range}</c>, awarding proportional credit
    /// equal to sequenceMatches / totalCells.
    /// </returns>
    /// <exception cref="Exception">Thrown when <c>rule.Range</c> is missing.</exception>
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

    /// <summary>
    /// Checks that non-blank cells within a range are numeric (ignores blanks).
    /// Awards proportional credit based on the fraction of non-blank numeric cells.
    /// </summary>
    /// <param name="rule">Requires <c>Range</c>. Honors <c>Points</c>.</param>
    /// <param name="wsS">Student worksheet.</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>range_numeric:{range}</c>. If the range contains only blanks,
    /// returns 0/pts with comment "no non-blank cells".
    /// </returns>
    /// <exception cref="Exception">Thrown when <c>rule.Range</c> is missing.</exception>
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
}
