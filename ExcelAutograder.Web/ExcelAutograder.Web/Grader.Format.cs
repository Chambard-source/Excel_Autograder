using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    private enum NumFmtKind
    {
        Unknown, General, Number, Currency, Accounting, Percent, Scientific, Fraction, Text,
        DateShort, DateLong, Time, DateTime, Custom
    }

    private static CheckResult GradeFormat(Rule rule, IXLWorksheet wsS)
    {
        if (string.IsNullOrWhiteSpace(rule.Cell))
        {
            if (!string.IsNullOrWhiteSpace(rule.Range))
                return GradeRangeFormat(rule, wsS);
            throw new Exception("format check missing 'cell'");
        }

        var cellAddr = rule.Cell!;
        var pts = rule.Points;
        var cell = wsS.Cell(cellAddr);

        (bool ok, string reason) OneFormat(RuleOption opt) =>
            FormatMatches(cell, opt.Format ?? rule.Format ?? new());

        var result = rule.AnyOf is { Count: > 0 }
            ? AnyOfMatch(rule.AnyOf, OneFormat)
            : OneFormat(new RuleOption());

        return new CheckResult($"format:{cellAddr}", pts, result.ok ? pts : 0, result.ok, result.reason);
    }

    private static (NumFmtKind Kind, int? Decimals, string Raw) AnalyzeNumberFormat(string? fmt)
    {
        var raw = (fmt ?? "").Trim();
        var f = Regex.Replace(raw, @"\[[^\]]*\]", "", RegexOptions.IgnoreCase);
        f = f.Split(';')[0].Trim();

        if (string.IsNullOrEmpty(f))
            return (NumFmtKind.General, null, raw);

        var lower = f.ToLowerInvariant();

        bool hasMonthName = Regex.IsMatch(lower, @"m{3,}");
        bool hasDayName = Regex.IsMatch(lower, @"d{3,}");
        bool hasDateParts = Regex.IsMatch(lower, @"\b[dmysh]\b|d|m|y", RegexOptions.IgnoreCase);
        bool hasTimeParts = lower.Contains("h") || lower.Contains("s") || lower.Contains("am/pm");

        if (lower == "general") return (NumFmtKind.General, null, raw);
        if (lower.Contains("%")) return (NumFmtKind.Percent, GetDecimalPlaces(lower), raw);
        if (lower.Contains("_(") || lower.Contains("* ") || (lower.Contains("€ ") && lower.Contains("_")))
            return (NumFmtKind.Accounting, GetDecimalPlaces(lower), raw);
        if (lower.Contains("$") || lower.Contains("¥") || lower.Contains("€") || lower.Contains("£") || lower.Contains("₩"))
            return (NumFmtKind.Currency, GetDecimalPlaces(lower), raw);
        if (hasDateParts && hasTimeParts) return (NumFmtKind.DateTime, null, raw);
        if (hasTimeParts) return (NumFmtKind.Time, null, raw);
        if (hasDateParts)
            return ((hasMonthName || hasDayName) ? NumFmtKind.DateLong : NumFmtKind.DateShort, null, raw);
        if (lower.Contains("e+")) return (NumFmtKind.Scientific, GetDecimalPlaces(lower), raw);
        if (lower.Contains("?/") || lower.Contains("#/")) return (NumFmtKind.Fraction, null, raw);
        if (lower.Contains("@")) return (NumFmtKind.Text, null, raw);

        if (Regex.IsMatch(lower, @"[#0](?:[#,]*[#0])?(?:\.(?<d>0+))?"))
            return (NumFmtKind.Number, GetDecimalPlaces(lower), raw);

        return (NumFmtKind.Custom, null, raw);
    }

    private static int? GetDecimalPlaces(string patternLower)
    {
        var m = Regex.Match(patternLower.Split(';')[0], @"\.(0+)");
        return m.Success ? m.Groups[1].Value.Length : (int?)null;
    }

    private static (bool ok, string reason) FormatMatches(IXLCell c, FormatSpec fmt)
    {
        var reasons = new List<string>();
        var style = c.Style;

        if (fmt.Font is not null)
        {
            if (fmt.Font.Name is not null && style.Font.FontName != fmt.Font.Name) reasons.Add("font name");
            if (fmt.Font.Size is not null && Math.Abs(style.Font.FontSize - fmt.Font.Size.Value) > 0.01) reasons.Add("font size");
            if (fmt.Font.Bold is not null && style.Font.Bold != fmt.Font.Bold.Value) reasons.Add("font bold");
            if (fmt.Font.Italic is not null && style.Font.Italic != fmt.Font.Italic.Value) reasons.Add("font italic");
        }
        if (fmt.Bold is not null && style.Font.Bold != fmt.Bold.Value) reasons.Add("bold");
        if (fmt.Italic is not null && style.Font.Italic != fmt.Italic.Value) reasons.Add("italic");

        if (fmt.NumberFormat is not null)
        {
            var want = AnalyzeNumberFormat(fmt.NumberFormat);
            var got = AnalyzeNumberFormat(style.NumberFormat.Format ?? "");

            if (want.Kind == NumFmtKind.Custom || want.Kind == NumFmtKind.Unknown ||
                got.Kind == NumFmtKind.Custom || got.Kind == NumFmtKind.Unknown)
            {
                if (!string.Equals(style.NumberFormat.Format ?? "", fmt.NumberFormat, StringComparison.Ordinal))
                    reasons.Add($"number_format literal ('{style.NumberFormat.Format ?? "General"}' != '{fmt.NumberFormat}')");
            }
            else
            {
                if (got.Kind != want.Kind)
                    reasons.Add($"number_format kind ({got.Kind.ToString().ToLower()} != {want.Kind.ToString().ToLower()})");

                bool decimalsMatter =
                    want.Kind is NumFmtKind.Number or NumFmtKind.Currency or NumFmtKind.Percent or NumFmtKind.Accounting;

                if (decimalsMatter && want.Decimals is not null)
                {
                    var gd = got.Decimals ?? 0;
                    var wd = want.Decimals.Value;
                    if (gd != wd)
                        reasons.Add($"decimals ({gd} != {wd})");
                }
            }
        }

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
}
