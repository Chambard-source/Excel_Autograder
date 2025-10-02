using System;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
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
}
