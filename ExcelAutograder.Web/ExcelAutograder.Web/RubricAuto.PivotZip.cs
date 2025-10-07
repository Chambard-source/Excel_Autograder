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

        // -- Load workbook + namespaces
        var wbEntry = zip.GetEntry("xl/workbook.xml");
        if (wbEntry is null) return map;

        var wbXml = XDocument.Load(wbEntry.Open());
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relns = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace odn = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // --- Build sheet index -> name
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

        // --- Build cacheId -> cache field names (from pivotCacheDefinition files)
        var cacheIdToNames = new Dictionary<int, string[]>();

        // workbook rels maps r:id -> target path
        var wbRelsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
        Dictionary<string, string> idToTarget = new();
        if (wbRelsEntry is not null)
        {
            var wbRels = XDocument.Load(wbRelsEntry.Open());
            idToTarget = wbRels.Root!
                .Elements(relns + "Relationship")
                .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/pivotCacheDefinition") == true)
                .ToDictionary(
                    r => (string)r.Attribute("Id")!,
                    r => ("xl/" + ((string)r.Attribute("Target")!).TrimStart('/')).Replace("../", "")
                );
        }

        // map cacheId -> r:id from workbook.xml
        var caches = wbXml.Root!.Element(ns + "pivotCaches")?.Elements(ns + "pivotCache") ?? Enumerable.Empty<XElement>();
        foreach (var pc in caches)
        {
            if (!int.TryParse((string?)pc.Attribute("cacheId"), out var cid)) continue;
            var rid = (string?)pc.Attribute(XName.Get("id", odn.NamespaceName));
            if (rid is null || !idToTarget.TryGetValue(rid, out var defPath)) continue;

            var defEntry = zip.GetEntry(defPath);
            if (defEntry is null) continue;

            var cdef = XDocument.Load(defEntry.Open());
            var names = cdef.Root!
                .Element(ns + "cacheFields")?
                .Elements(ns + "cacheField")
                .Select(cf => ((string?)cf.Attribute("name"))?.Trim() ?? "")
                .ToArray() ?? Array.Empty<string>();

            cacheIdToNames[cid] = names;
        }

        // --- For each sheet, find pivotTables and emit rules
        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsEntry = zip.GetEntry(relsPath);
            if (relsEntry is null) continue;

            var relsXml = XDocument.Load(relsEntry.Open());
            var relTargets = relsXml.Root!.Elements(relns + "Relationship")
                .Where(e => ((string?)e.Attribute("Type"))?.Contains("pivotTable", StringComparison.OrdinalIgnoreCase) == true)
                .Select(e => (string?)e.Attribute("Target"))
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList();

            foreach (var target in relTargets)
            {
                var full = target!.StartsWith("/")
                             ? target.TrimStart('/')
                             : target.StartsWith("../")
                                ? "xl/" + target.Replace("../", "")
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

                // cacheId from pivotTableDefinition -> look up cache field names
                int cacheId = 0;
                int.TryParse((string?)def.Attribute("cacheId"), out cacheId);
                cacheIdToNames.TryGetValue(cacheId, out var cacheFieldNames);

                // pivotFields (some files have @name here; many leave it empty)
                var pivotFields = def.Element(ns + "pivotFields")?.Elements(ns + "pivotField").ToList()
                                   ?? new List<XElement>();

                string NameByIndex(int fi)
                {
                    // 1) pivotField @name if present
                    if (fi >= 0 && fi < pivotFields.Count)
                    {
                        var pf = pivotFields[fi];
                        var nAttr = ((string?)pf.Attribute("name"))?.Trim();
                        if (!string.IsNullOrWhiteSpace(nAttr)) return nAttr!;
                    }
                    // 2) cache field name (real source column)
                    if (cacheFieldNames is not null && fi >= 0 && fi < cacheFieldNames.Length)
                    {
                        var n = cacheFieldNames[fi]?.Trim();
                        if (!string.IsNullOrWhiteSpace(n)) return n!;
                    }
                    // 3) fallback
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

                // Values: include caption + normalized agg + source field name (via fld index)
                var values = new List<PivotValueSpec>();
                foreach (var df in def.Element(ns + "dataFields")?.Elements(ns + "dataField") ?? Enumerable.Empty<XElement>())
                {
                    var caption = ((string?)df.Attribute("name"))?.Trim();
                    var fldAttr = (string?)df.Attribute("fld");
                    var src = (int.TryParse(fldAttr, out var fi)) ? NameByIndex(fi) : null;

                    var subtotal = ((string?)df.Attribute("subtotal"))?.ToLowerInvariant() ?? "sum";
                    string agg = subtotal.Contains("count") ? "count"
                               : subtotal.Contains("average") || subtotal.Contains("avg") ? "average"
                               : subtotal.Contains("min") ? "min"
                               : subtotal.Contains("max") ? "max"
                               : "sum";

                    // prefer caption if present; else synthesize "Sum of X"
                    if (string.IsNullOrWhiteSpace(caption) && !string.IsNullOrWhiteSpace(src))
                        caption = $"{char.ToUpper(agg[0]) + agg[1..]} of {src}";

                    values.Add(new PivotValueSpec
                    {
                        Field = caption ?? "Value",
                        Agg = agg,
                        Source = src   // <-- add this property to your PivotValueSpec (string? Source)
                    });
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
