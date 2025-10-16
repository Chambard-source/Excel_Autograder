using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

public static partial class RubricAuto
{
    /// <summary>
    /// Scans the key workbook’s XLSX ZIP for conditional formatting definitions and
    /// emits rubric rules of type <c>"conditional_format"</c>, grouped by sheet name.
    /// </summary>
    /// <param name="keyZipBytes">Raw XLSX (ZIP) bytes for the key workbook.</param>
    /// <returns>
    /// A dictionary mapping sheet name → list of <see cref="Rule"/> objects describing
    /// conditional formatting expectations (range, type, operator, formulas, text, fill color).
    /// </returns>
    /// <remarks>
    /// This inspects <c>xl/workbook.xml</c> to enumerate sheets, then for each sheet reads
    /// <c>xl/worksheets/sheetN.xml</c> and its <c>&lt;conditionalFormatting&gt;</c> blocks.
    /// If present, <c>xl/styles.xml</c> is used to resolve <c>dxfId</c> to an RGB fill (alpha stripped).
    /// Only the first range in <c>sqref</c> is captured for each rule.
    /// </remarks>
    internal static Dictionary<string, List<Rule>> ExtractConditionalRulesFromZip(byte[] keyZipBytes)
    {
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(keyZipBytes);
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        var wbEntry = zip.GetEntry("xl/workbook.xml"); if (wbEntry is null) return map;
        var wbXml = System.Xml.Linq.XDocument.Load(wbEntry.Open());
        XName S(string n) => System.Xml.Linq.XName.Get(n, "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        var idxToName = new Dictionary<int, string>(); int idx = 1;
        var sheetsEl = wbXml.Root?.Element(S("sheets"));
        if (sheetsEl != null)
            foreach (var sh in sheetsEl.Elements(S("sheet")))
                idxToName[idx++] = (string?)sh.Attribute("name") ?? $"Sheet{idx}";

        // styles for dxf (optional)
        var stylesEntry = zip.GetEntry("xl/styles.xml");
        System.Xml.Linq.XDocument? stylesXml = stylesEntry != null ? System.Xml.Linq.XDocument.Load(stylesEntry.Open()) : null;

        for (int i = 1; i < idx; i++)
        {
            if (!idxToName.TryGetValue(i, out var sheetName)) continue;
            var sheetPath = $"xl/worksheets/sheet{i}.xml";
            var sheetEntry = zip.GetEntry(sheetPath); if (sheetEntry is null) continue;
            var wsXml = System.Xml.Linq.XDocument.Load(sheetEntry.Open());

            foreach (var block in wsXml.Root!.Elements(S("conditionalFormatting")))
            {
                var sqref = (string?)block.Attribute("sqref") ?? "";
                foreach (var ruleEl in block.Elements(S("cfRule")))
                {
                    var t = (string?)ruleEl.Attribute("type");
                    var op = (string?)ruleEl.Attribute("operator");
                    var frms = ruleEl.Elements(S("formula")).Select(e => e.Value?.Trim()).ToList();
                    var txt = (string?)ruleEl.Attribute("text");

                    // Resolve fill color from styles (if a dxf is referenced)
                    string? fillRgb = null;
                    var dxfIdAttr = ruleEl.Attribute("dxfId");
                    if (stylesXml != null && dxfIdAttr != null && int.TryParse(dxfIdAttr.Value, out var dxfId))
                    {
                        var dxfs = stylesXml.Root?.Element(S("dxfs"))?.Elements(S("dxf")).ToList();
                        if (dxfs != null && dxfId >= 0 && dxfId < dxfs.Count)
                        {
                            var dxf = dxfs[dxfId];
                            fillRgb = dxf.Element(S("fill"))?.Element(S("patternFill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value
                                   ?? dxf.Element(S("fill"))?.Element(S("fgColor"))?.Attribute("rgb")?.Value;
                            if (!string.IsNullOrWhiteSpace(fillRgb) && fillRgb.Length == 8) // strip ARGB alpha (FFRRGGBB → RRGGBB)
                                fillRgb = fillRgb.Substring(2);
                        }
                    }

                    var rule = new Rule
                    {
                        Type = "conditional_format",
                        Points = 0.5, // will be rescaled later
                        Note = "Conditional format",
                        Cond = new ConditionalFormatSpec
                        {
                            Sheet = sheetName,
                            Range = sqref.Split(' ').FirstOrDefault(), // first target range
                            Type = t,
                            Operator = MapXmlOp(op),
                            Formula1 = frms.ElementAtOrDefault(0),
                            Formula2 = frms.ElementAtOrDefault(1),
                            Text = txt,
                            FillRgb = fillRgb
                        }
                    };

                    if (!map.TryGetValue(sheetName, out var list)) map[sheetName] = list = new List<Rule>();
                    list.Add(rule);
                }
            }
        }
        return map;
    }

    /// <summary>
    /// Maps SpreadsheetML operator tokens (e.g., <c>greaterThan</c>) to the compact tokens
    /// the grader uses (e.g., <c>gt</c>).
    /// </summary>
    /// <param name="op">Raw operator string from <c>&lt;cfRule operator="…"/&gt;</c>.</param>
    /// <returns>Compact operator code, or the original token if unknown.</returns>
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
}
