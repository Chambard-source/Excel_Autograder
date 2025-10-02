using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

public static partial class RubricAuto
{
    
    /// <summary>
    /// Parse the KEY workbook ZIP to discover pivot tables and emit pivot_layout rules
    /// keyed by sheet name.
    /// </summary>
    internal static Dictionary<string, List<Rule>> ExtractPivotRulesFromZip(byte[] keyZipBytes)
    {
        // sheetName -> rules
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(keyZipBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);

        // Map sheet index to sheet name from xl/workbook.xml
        var wbEntry = zip.GetEntry("xl/workbook.xml");
        if (wbEntry is null) return map;

        var wbXml = XDocument.Load(wbEntry.Open());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        var sheetIndexToName = new Dictionary<int, string>();
        int idx = 1;
        var sheetsEl = wbXml.Root?.Element(ns + "sheets");
        if (sheetsEl != null)
        {
            foreach (var sh in sheetsEl.Elements(ns + "sheet"))
            {
                var n = (string?)sh.Attribute("name") ?? $"Sheet{idx}";
                sheetIndexToName[idx] = n;
                idx++;
            }
        }

        // For each sheetN.xml, open its _rels to find pivotTable targets
        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsEntry = zip.GetEntry(relsPath);
            if (relsEntry is null) continue;

            var relsXml = XDocument.Load(relsEntry.Open());
            var rels = relsXml.Root!.Elements()
                        .Where(e => ((string?)e.Attribute("Type"))?.Contains("pivotTable", StringComparison.OrdinalIgnoreCase) == true)
                        .Select(e => (string?)e.Attribute("Target"))
                        .Where(t => !string.IsNullOrWhiteSpace(t))
                        .ToList();


            foreach (var target in rels)
            {
                var normalized = target!.Replace("\\", "/");
                var full = target.StartsWith("/") ? target.TrimStart('/')
                         : target.StartsWith("../") ? "xl/" + target.Replace("../", "")
                         : "xl/worksheets/" + target;

                if (full.Contains("/pivotTables/", StringComparison.OrdinalIgnoreCase))
                {
                    var pos = full.IndexOf("/pivotTables/", StringComparison.OrdinalIgnoreCase);
                    full = "xl" + full.Substring(pos);
                }

                var ptEntry = zip.GetEntry(full);
                if (ptEntry is null) continue;

                var ptXml = XDocument.Load(ptEntry.Open());
                var def = ptXml.Root!;
                var ptName = (string?)def.Attribute("name") ?? "Pivot";

                // Gather pivotFields to map indexes → names
                var pivotFields = def.Element(ns + "pivotFields")?.Elements(ns + "pivotField").ToList()
                                   ?? new List<XElement>();

                string NameByIndex(int fi)
                {
                    if (fi >= 0 && fi < pivotFields.Count)
                    {
                        var pf = pivotFields[fi];
                        var nAttr = (string?)pf.Attribute("name");
                        if (!string.IsNullOrWhiteSpace(nAttr)) return nAttr!;
                    }
                    return $"Field{fi}";
                }

                var rows = new List<string>();
                foreach (var rf in def.Element(ns + "rowFields")?.Elements(ns + "field") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)rf.Attribute("x"), out var fi)) rows.Add(NameByIndex(fi));

                var cols = new List<string>();
                foreach (var cf in def.Element(ns + "colFields")?.Elements(ns + "field") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)cf.Attribute("x"), out var fi)) cols.Add(NameByIndex(fi));

                var filters = new List<string>();
                foreach (var pf in def.Element(ns + "pageFields")?.Elements(ns + "pageField") ?? Enumerable.Empty<XElement>())
                    if (int.TryParse((string?)pf.Attribute("fld"), out var fi)) filters.Add(NameByIndex(fi));

                var values = new List<PivotValueSpec>();
                foreach (var df in def.Element(ns + "dataFields")?.Elements(ns + "dataField") ?? Enumerable.Empty<XElement>())
                {
                    var nm = (string?)df.Attribute("name");
                    var fld = (string?)df.Attribute("fld");
                    string fieldName = !string.IsNullOrWhiteSpace(nm) ? nm!
                                      : (int.TryParse(fld, out var fi) ? NameByIndex(fi) : "Value");

                    var subtotal = ((string?)df.Attribute("subtotal"))?.ToLowerInvariant() ?? "sum";
                    string agg = subtotal.Contains("count") ? "count"
                               : subtotal.Contains("average") || subtotal.Contains("avg") ? "average"
                               : subtotal.Contains("min") ? "min"
                               : subtotal.Contains("max") ? "max"
                               : "sum";

                    values.Add(new PivotValueSpec { Field = fieldName, Agg = agg });
                }

                var rule = new Rule
                {
                    Type = "pivot_layout",
                    Points = 1.5,
                    Note = $"Pivot '{ptName}' layout",
                    Pivot = new PivotSpec
                    {
                        Sheet = sheetName,
                        TableNameLike = ptName,
                        Rows = rows.Count > 0 ? rows.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Columns = cols.Count > 0 ? cols.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Filters = filters.Count > 0 ? filters.Distinct(StringComparer.OrdinalIgnoreCase).ToList() : null,
                        Values = values.Count > 0 ? values : null
                    }
                };

                if (!map.TryGetValue(sheetName, out var list)) map[sheetName] = list = new List<Rule>();
                list.Add(rule);
            }
        }

        return map;
    }
}
