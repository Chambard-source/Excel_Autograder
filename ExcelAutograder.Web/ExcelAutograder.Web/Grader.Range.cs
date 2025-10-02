using System;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
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
}
