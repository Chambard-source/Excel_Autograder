using System;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Grades the value of a target cell against an expected value.
    /// Supports (in priority order): expected-from-key, regex match, explicit option expected, rule-level expected.
    /// Numeric values are compared within an absolute tolerance; otherwise string equality is used
    /// with optional case sensitivity.
    /// </summary>
    /// <param name="rule">
    /// Rule specifying:
    /// <list type="bullet">
    ///   <item><description><c>Cell</c> (A1 address)</description></item>
    ///   <item><description><c>Points</c>, <c>Tolerance</c></description></item>
    ///   <item><description><c>ExpectedFromKey</c> or <c>Expected</c> or <c>AnyOf[].Expected / .ExpectedRegex</c></description></item>
    ///   <item><description><c>CaseSensitive</c> (optional)</description></item>
    /// </list>
    /// </param>
    /// <param name="wsS">Student worksheet.</param>
    /// <param name="wsK">Key worksheet (used when <c>ExpectedFromKey</c> is true).</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>value:{cell}</c>, full points if matched, else 0 with reason text.
    /// </returns>
    /// <exception cref="Exception">Thrown if <c>rule.Cell</c> is missing.</exception>
    private static CheckResult GradeValue(Rule rule, IXLWorksheet wsS, IXLWorksheet? wsK)
    {
        var cellAddr = rule.Cell ?? throw new Exception("value check missing 'cell'");
        var pts = rule.Points;
        var tol = rule.Tolerance ?? 0.0;

        // Evaluates a single RuleOption path (regex / expected literal / expected from key).
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
