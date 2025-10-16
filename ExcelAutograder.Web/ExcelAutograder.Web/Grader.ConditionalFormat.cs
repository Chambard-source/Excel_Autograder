using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Maps OOXML conditional format operator tokens (e.g., <c>greaterThan</c>)
    /// to the shorter rubric/operator codes used in your JSON (<c>gt</c>, <c>le</c>, etc.).
    /// </summary>
    /// <param name="op">The OOXML operator token from <c>cfRule@operator</c>.</param>
    /// <returns>The normalized short token (e.g., <c>gt</c>), or the original string when unknown.</returns>
    private static string? MapXmlOp(string? op) => op switch
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

    /// <summary>
    /// Grades whether a specific Conditional Formatting rule exists on the student workbook.
    /// Inspects raw OOXML (via cached student .xlsx bytes) to avoid ClosedXML limitations,
    /// and compares against <see cref="Rule.Cond"/> (type, operator, formulas, text, fill color, range).
    /// </summary>
    /// <param name="rule">The rubric rule containing a <see cref="ConditionalFormatSpec"/>.</param>
    /// <param name="wbS">Student workbook (used to seed cached .xlsx bytes).</param>
    /// <returns>
    /// <see cref="CheckResult"/> with id <c>conditional_format:{sheet}</c>, full points if a matching CF is found,
    /// otherwise 0 with an explanation of what was expected and what was seen.
    /// </returns>
    private static CheckResult GradeConditionalFormat(Rule rule, XLWorkbook wbS)
    {
        EnsureStudentZipBytes(wbS);
        var pts = rule.Points;
        var spec = rule.Cond ?? new ConditionalFormatSpec();
        var sheetName = spec.Sheet ?? wbS.Worksheets.First().Name;

        if (_zipBytes.Value is null)
            return new CheckResult("conditional_format", pts, 0, false,
                "Student .xlsx bytes not available to inspect conditional formats");

        var expected = DescribeCond(spec);
        var ok = FindCFInStudentZip_NoClosedXml(_zipBytes.Value, sheetName, spec,
                                                out var reason, out var matchedSummary);

        var note = ok
            ? $"matched: {matchedSummary}"
            : $"expected: {expected}; {reason}";

        return new CheckResult($"conditional_format:{sheetName}", pts, ok ? pts : 0, ok, note);
    }

    /// <summary>
    /// Searches the student's raw <c>.xlsx</c> (ZIP) for a conditional formatting rule
    /// matching the given <see cref="ConditionalFormatSpec"/> without relying on ClosedXML’s CF APIs.
    /// </summary>
    /// <param name="zipBytes">Raw bytes of the student workbook.</param>
    /// <param name="sheetName">Target sheet name.</param>
    /// <param name="spec">Expected conditional-format attributes to match.</param>
    /// <param name="reason">On failure, a human-friendly explanation of why matching failed.</param>
    /// <param name="matchedSummary">On success, a short diagnostic summary of the matched rule.</param>
    /// <returns><c>true</c> if a matching rule is found; otherwise <c>false</c> with <paramref name="reason"/>.</returns>
    /// <remarks>
    /// Parses:
    /// <list type="bullet">
    ///   <item><description><c>xl/workbook.xml</c> to map sheet names → indices</description></item>
    ///   <item><description><c>xl/worksheets/sheetN.xml</c> to enumerate <c>&lt;conditionalFormatting&gt;</c> blocks</description></item>
    ///   <item><description><c>xl/styles.xml</c> to resolve <c>dxfId</c> → fill color (ARGB)</description></item>
    /// </list>
    /// Matching rules:
    /// <list type="bullet">
    ///   <item><description>Sheet &amp; range overlap (via <c>sqref</c>)</description></item>
    ///   <item><description>Type (<c>rule.Type</c>) and operator (mapped via <see cref="MapXmlOp"/>)</description></item>
    ///   <item><description>Formulas 1/2 (normalized: strip leading '=' and spaces)</description></item>
    ///   <item><description>Contains-text for <c>text</c> attribute (when provided)</description></item>
    ///   <item><description>Fill RGB suffix match (DXF ARGB → RGB, case-insensitive)</description></item>
    /// </list>
    /// </remarks>
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

            // Normalizes a formula string for comparison (trim, drop '=', remove spaces).
            static string Norm(string? s)
            {
                if (string.IsNullOrWhiteSpace(s)) return "";
                s = s.Trim();
                if (s.StartsWith("=")) s = s.Substring(1);
                return s.Replace(" ", "");
            }

            // Resolve ARGB → RGB fill color from styles.xml dxfs[dxfId].
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

            // True if any sqref token intersects the expected A1 range (when provided).
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
                        // Excel often stores ARGB; we strip alpha. Compare suffix to allow theme/alpha differences.
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
}
