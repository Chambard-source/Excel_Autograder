using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    // Move the existing implementations here from your original Grader.cs:

    private static string Normalize(object? o) => (o?.ToString() ?? "").Trim();
    private static string NormalizeFormula(string? f)
    {
        var s = (f ?? "").Trim();
        if (s.Length == 0) return "";
        if (s[0] != '=') s = "=" + s;
        s = s.Replace(" ", "").Replace("$", "");
        return s.ToUpperInvariant();
    }

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
    private static (bool ok, string detail) ValueMatches(IXLCell cell, Rule r)
    {
        var expectedLiteral = GetExpectedLiteral(r);     // ← normalized string/number/bool
        var expectedRegex = r.ExpectedRegex;

        // nothing to check
        if (string.IsNullOrWhiteSpace(expectedRegex) && string.IsNullOrWhiteSpace(expectedLiteral))
            return (true, "");

        // what the grader "sees"
        var studentText = cell.GetFormattedString();
        if (string.IsNullOrWhiteSpace(studentText))
            studentText = cell.GetString(); // fallback to raw

        // regex has priority
        if (!string.IsNullOrWhiteSpace(expectedRegex))
        {
            bool pass = Regex.IsMatch(studentText, expectedRegex!, RegexOptions.IgnoreCase);
            return pass
                ? (true, "value matches regex")
                : (false, $"value '{studentText}' !~ /{expectedRegex}/");
        }

        // otherwise compare against expected literal (numeric tolerant)
        var want = expectedLiteral ?? "";

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

        if (TryParseNumber(studentText, out var gotNum) && TryParseNumber(want, out var expNum))
        {
            var tol = r.Tolerance ?? 0.0;
            bool pass = Math.Abs(gotNum - expNum) <= Math.Abs(tol);
            return pass
                ? (true, $"value {gotNum}≈{expNum} (±{tol})")
                : (false, $"value {gotNum}≠{expNum} (tol {tol})");
        }

        // string compare (case-insensitive)
        bool textPass = string.Equals(studentText.Trim(), want.Trim(), StringComparison.OrdinalIgnoreCase);
        return textPass
            ? (true, "value text matches")
            : (false, $"value text got '{studentText}' expected '{want}'");
    }
    private static string? GetExpectedLiteral(Rule rr)
    {
        if (rr.Expected is null) return null;

        if (rr.Expected is System.Text.Json.JsonElement je)
        {
            return je.ValueKind switch
            {
                System.Text.Json.JsonValueKind.String => je.GetString(),
                System.Text.Json.JsonValueKind.Number =>
                    je.TryGetDouble(out var d)
                        ? d.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        : je.ToString(),
                System.Text.Json.JsonValueKind.True => "TRUE",
                System.Text.Json.JsonValueKind.False => "FALSE",
                System.Text.Json.JsonValueKind.Null => null,
                _ => je.ToString()
            };
        }

        return rr.Expected.ToString();
    }
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

    // Color helpers: --------------------------------------------------------------------------------------
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
    private static string XLColorToArgb(XLColor color)
    {
        // ClosedXML 0.105.x: XLColor.Color is System.Drawing.Color
        var sys = color.Color;
        return $"FF{sys.R:X2}{sys.G:X2}{sys.B:X2}";
    }
    
    private static bool ArgbEqual(string a, string b) => NormalizeArgb(a) == NormalizeArgb(b);

    
    // A1 helpers used by CF/table (if they existed in your file): -------------------------------------------
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
    private static bool RectsIntersect((int r1, int c1, int r2, int c2) a, (int r1, int c1, int r2, int c2) b)
    {
        return !(b.c1 > a.c2 || b.c2 < a.c1 || b.r1 > a.r2 || b.r2 < a.r1);
    }

    // CF description helpers if present: ---------------------------------------------------------------------
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
}
