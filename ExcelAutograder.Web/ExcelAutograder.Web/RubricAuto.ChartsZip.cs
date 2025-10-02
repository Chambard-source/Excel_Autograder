using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

public static partial class RubricAuto
{
    /// <summary>
    /// Parse the KEY workbook ZIP to discover charts and emit rules
    /// (type "chart") keyed by sheet name.
    /// </summary>
    internal static Dictionary<string, List<Rule>> ExtractChartRulesFromZip(byte[] keyZipBytes)
    {
        var map = new Dictionary<string, List<Rule>>(StringComparer.OrdinalIgnoreCase);
        var chartsBySheet = ParseChartsFromZipAuto(keyZipBytes);

        foreach (var kv in chartsBySheet)
        {
            var sheet = kv.Key;
            foreach (var ch in kv.Value)
            {
                var rule = new Rule
                {
                    Type = "chart",
                    Points = 1.5, // UI rescaled later
                    Note = $"Chart '{ch.Name}' on {sheet}",
                    Chart = new ChartSpec
                    {
                        Sheet = sheet,
                        NameLike = ch.Name,
                        Type = ch.Type,
                        Title = ch.Title,
                        TitleRef = ch.TitleRef,
                        LegendPos = ch.LegendPos,
                        DataLabels = ch.DataLabels,
                        XTitle = ch.XTitle,
                        YTitle = ch.YTitle,
                        Series = ch.Series.Select(s => new ChartSeriesSpec
                        {
                            Name = s.Name,
                            NameRef = s.NameRef,
                            CatRef = s.CatRef,
                            ValRef = s.ValRef
                        }).ToList()
                    }
                };

                if (!map.TryGetValue(sheet, out var list)) map[sheet] = list = new();
                list.Add(rule);
            }
        }

        return map;
    }

    private static Dictionary<string, List<AutoChartInfo>> ParseChartsFromZipAuto(byte[] zipBytes)
    {
        var result = new Dictionary<string, List<AutoChartInfo>>(StringComparer.OrdinalIgnoreCase);

        using var ms = new MemoryStream(zipBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);

        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace pkg = "http://schemas.openxmlformats.org/package/2006/relationships";

        string ReadEntryText(string path)
        {
            var e = zip.GetEntry(path);
            if (e == null) return "";
            using var s = e.Open();
            using var r = new StreamReader(s);
            return r.ReadToEnd();
        }

        // 1) map sheet index -> sheet name (sheet1.xml → "Summary", etc.)
        var wbTxt = ReadEntryText("xl/workbook.xml");
        if (string.IsNullOrEmpty(wbTxt)) return result;
        var wb = XDocument.Parse(wbTxt);

        var sheetIndexToName = new Dictionary<int, string>();
        var sheetsEl = wb.Root?.Element(nsMain + "sheets");
        int idx = 1;
        if (sheetsEl != null)
        {
            foreach (var sh in sheetsEl.Elements(nsMain + "sheet"))
            {
                var nm = (string?)sh.Attribute("name") ?? $"Sheet{idx}";
                sheetIndexToName[idx++] = nm;
            }
        }

        // 2) for each sheet rels → drawing → chart
        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsTxt = ReadEntryText(relsPath);
            if (string.IsNullOrEmpty(relsTxt)) continue;

            var rels = XDocument.Parse(relsTxt);
            var drawingTargets = rels.Root?
                .Elements(pkg + "Relationship")  // <— was rel + "Relationship"
                .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/drawing") == true)
                .Select(r => ((string?)r.Attribute("Target"))?.TrimStart('/').Replace("../", "xl/"))
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .ToList() ?? new List<string>();

            foreach (var drawingTarget in drawingTargets)
            {
                var drPath = drawingTarget!.StartsWith("xl/") ? drawingTarget : $"xl/{drawingTarget}";
                var drTxt = ReadEntryText(drPath);
                if (string.IsNullOrEmpty(drTxt)) continue;
                var drXml = XDocument.Parse(drTxt);

                // drawing rels
                var drRelsPath = drPath.Replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
                var drRelsTxt = ReadEntryText(drRelsPath);
                if (string.IsNullOrEmpty(drRelsTxt)) continue;
                var drRels = XDocument.Parse(drRelsTxt);

                // find c:chart r:id within graphicFrames
                var frames = drXml.Descendants(xdr + "graphicFrame").ToList();
                foreach (var gf in frames)
                {
                    var cNvPr = gf.Element(xdr + "nvGraphicFramePr")?.Element(xdr + "cNvPr");
                    var frameName = cNvPr?.Attribute("name")?.Value ?? "Chart";

                    var chartElem = gf.Descendants(a + "graphicData").Descendants(c + "chart").FirstOrDefault();
                    var rid = chartElem?.Attribute(rel + "id")?.Value;
                    if (string.IsNullOrWhiteSpace(rid)) continue;

                    var target = drRels.Root?
                        .Elements(pkg + "Relationship")  // <— was rel + "Relationship"
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;

                    if (string.IsNullOrWhiteSpace(target)) continue;

                    var tgt = (target ?? "").Replace("\\", "/");
                    var chartPath = tgt.StartsWith("/") ? tgt.TrimStart('/')
                                 : tgt.StartsWith("../") ? "xl/" + tgt.Substring(3)
                                 : tgt.StartsWith("xl/") ? tgt : "xl/" + tgt;

                    var chTxt = ReadEntryText(chartPath);
                    if (string.IsNullOrEmpty(chTxt)) continue;

                    var chDoc = XDocument.Parse(chTxt);
                    var info = new AutoChartInfo { Sheet = sheetName, Name = frameName };

                    var plotArea = chDoc.Descendants(c + "plotArea").FirstOrDefault();
                    info.Type = DetectChartType(plotArea, c);   // or DetectChartType(plotArea) if your helper takes 1 arg

                    // Title
                    var titleEl = chDoc.Descendants(c + "title").FirstOrDefault();
                    (info.Title, info.TitleRef) = ReadChartTextAuto(titleEl, c, a);

                    // Axis titles
                    var catAx = plotArea?.Elements(c + "catAx").FirstOrDefault();
                    var valAx = plotArea?.Elements(c + "valAx").FirstOrDefault();
                    (info.XTitle, _) = ReadChartTextAuto(catAx?.Element(c + "title"), c, a);
                    (info.YTitle, _) = ReadChartTextAuto(valAx?.Element(c + "title"), c, a);

                    // Legend + labels
                    var leg = chDoc.Descendants(c + "legend").FirstOrDefault();
                    info.LegendPos = leg?.Element(c + "legendPos")?.Attribute("val")?.Value;
                    info.DataLabels = plotArea?.Descendants(c + "dLbls").Any() == true;

                    // Series
                    foreach (var ser in plotArea?.Descendants().Where(e => e.Name.LocalName == "ser") ?? Enumerable.Empty<XElement>())
                    {
                        var si = new AutoSeriesInfo();
                        var tx = ser.Element(c + "tx");
                        (si.Name, si.NameRef) = ReadChartTextAuto(tx, c, a);

                        var cat = ser.Element(c + "cat");
                        si.CatRef = cat?.Element(c + "strRef")?.Element(c + "f")?.Value
                                    ?? cat?.Element(c + "numRef")?.Element(c + "f")?.Value;

                        var val = ser.Element(c + "val");
                        si.ValRef = val?.Element(c + "numRef")?.Element(c + "f")?.Value
                                    ?? val?.Element(c + "strRef")?.Element(c + "f")?.Value;

                        info.Series.Add(si);
                    }

                    if (!result.TryGetValue(sheetName, out var list)) result[sheetName] = list = new();
                    list.Add(info);
                }
            }
        }

        return result;

        // ---- helpers ----
        static (string? txt, string? cellRef) ReadChartTextAuto(XElement? node, XNamespace cns, XNamespace ans)
        {
            if (node == null) return (null, null);
            var tx = node.Element(cns + "tx");
            if (tx == null) return (null, null);

            var rich = tx.Element(cns + "rich");
            if (rich != null)
            {
                var text = string.Join("", rich.Descendants(ans + "t").Select(t => t.Value));
                return (text, null);
            }
            var strRef = tx.Element(cns + "strRef");
            var f = strRef?.Element(cns + "f")?.Value;
            return (null, f);
        }
    }

    private static string DetectChartType(System.Xml.Linq.XElement? plotArea, System.Xml.Linq.XNamespace c)
    {
        if (plotArea == null) return "";
        if (plotArea.Element(c + "bar3DChart") != null)
            return string.Equals(plotArea.Element(c + "bar3DChart")?.Element(c + "barDir")?.Attribute("val")?.Value, "col", StringComparison.OrdinalIgnoreCase) ? "column3D" : "bar3D";
        var bar = plotArea.Element(c + "barChart");
        if (bar != null)
            return string.Equals(bar.Element(c + "barDir")?.Attribute("val")?.Value, "col", StringComparison.OrdinalIgnoreCase) ? "column" : "bar";
        if (plotArea.Element(c + "line3DChart") != null) return "line3D";
        if (plotArea.Element(c + "area3DChart") != null) return "area3D";
        if (plotArea.Element(c + "pie3DChart") != null) return "pie3D";
        if (plotArea.Element(c + "pieChart") != null) return "pie";
        if (plotArea.Element(c + "areaChart") != null) return "area";
        if (plotArea.Element(c + "lineChart") != null) return "line";
        if (plotArea.Element(c + "scatterChart") != null) return "scatter";
        if (plotArea.Element(c + "bubbleChart") != null) return "bubble";
        if (plotArea.Element(c + "radarChart") != null) return "radar";
        if (plotArea.Element(c + "stockChart") != null) return "stock";
        return "";
    }


    // --------------------- CHART RULES (auto from key) ---------------------
    private class AutoChartInfo
    {
        public string Sheet = "";
        public string Name = "Chart";
        public string Type = ""; // line/column/bar/pie/scatter/area/doughnut/radar/bubble
        public string? Title, TitleRef;
        public string? LegendPos;
        public bool DataLabels;
        public string? XTitle, YTitle;
        public List<AutoSeriesInfo> Series = new();
    }
    private class AutoSeriesInfo
    {
        public string? Name, NameRef, CatRef, ValRef;
    }
}
