using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using ClosedXML.Excel;

public static partial class Grader
{
    /// <summary>
    /// Parsed chart metadata extracted from OOXML.
    /// </summary>
    private class ChartInfo
    {
        /// <summary>Worksheet the chart belongs to.</summary>
        public string Sheet = "";
        /// <summary>Frame name (e.g., "Chart 1").</summary>
        public string Name = "";
        /// <summary>Normalized chart type (line/column/bar/pie/scatter/area/doughnut/radar/bubble/pie3D).</summary>
        public string Type = "";
        /// <summary>Chart title text (if rich text), otherwise <c>null</c>.</summary>
        public string? Title, TitleRef;
        /// <summary>Legend position token from OOXML (e.g., right, left, top).</summary>
        public string? LegendPos;
        /// <summary>True if any <c>&lt;c:dLbls&gt;</c> block is present.</summary>
        public bool DataLabels;
        /// <summary>Category (X) axis title text.</summary>
        public string? XTitle;
        /// <summary>Value (Y) axis title text.</summary>
        public string? YTitle;
        /// <summary>Series list captured from plot area.</summary>
        public List<SeriesInfo> Series = new();
    }

    /// <summary>
    /// Parsed series metadata for a chart.
    /// </summary>
    private class SeriesInfo
    {
        /// <summary>Series name literal (if rich text).</summary>
        public string? Name;
        /// <summary>Series name cell reference (A1) when title comes from a cell.</summary>
        public string? NameRef;
        /// <summary>Categories reference (A1) – <c>strRef</c> or <c>numRef</c>.</summary>
        public string? CatRef;
        /// <summary>Values reference (A1) – usually <c>numRef</c>.</summary>
        public string? ValRef;
    }

    /// <summary>
    /// Grades a chart in the student workbook against a <see cref="ChartSpec"/>.
    /// The check builds a section-aware ID, parses all charts from the student's OOXML,
    /// filters by sheet (if provided), and then scores the best-matching chart by comparing:
    /// name-like, type, title/titleRef, legend position, data-labels presence, axis titles, and series refs.
    /// </summary>
    /// <param name="rule">Rule containing <see cref="Rule.Chart"/> expectations.</param>
    /// <param name="wbS">Student workbook (used to capture OOXML bytes).</param>
    /// <returns>
    /// <see cref="CheckResult"/> awarding a fraction of points proportional to matched checks,
    /// with a concise summary of missed attributes when not perfect.
    /// </returns>
    private static CheckResult GradeChart(Rule rule, XLWorkbook wbS)
    {
        var pts = rule.Points;
        var spec = rule.Chart;

        // Section-aware id that matches the UI grouping rules
        string SectionIdForSpec()
        {
            var sh = spec?.Sheet ?? "";
            var nm = spec?.NameLike ?? "";

            if (!string.IsNullOrWhiteSpace(nm))
                return $"chart:{sh}{(string.IsNullOrWhiteSpace(sh) ? "" : "/")}{nm}";

            // prefer TYPE first so column/pie/pie3D separate even if titles are present
            var typ = spec?.Type ?? "";
            if (!string.IsNullOrWhiteSpace(sh) && !string.IsNullOrWhiteSpace(typ))
                return $"chart:{sh}|type={typ.ToLowerInvariant()}";

            // if we fall back to TITLE, include the value to avoid collisions
            static string NormTitle(string s) =>
                System.Text.RegularExpressions.Regex.Replace((s ?? "").Trim(), @"\s+", " ").ToLowerInvariant();

            var title = spec?.Title ?? "";
            if (!string.IsNullOrWhiteSpace(sh) && !string.IsNullOrWhiteSpace(title))
                return $"chart:{sh}|title={NormTitle(title)}";

            // final fallback: series signature (first ValRef), normalized
            static string NormRefLocal(string? a1)
            {
                var s = (a1 ?? "").Replace("$", "").ToUpperInvariant();
                int bang = s.IndexOf('!');
                return bang >= 0 ? s[(bang + 1)..] : s;
            }
            var sig = spec?.Series is { Count: > 0 } ? NormRefLocal(spec.Series[0].ValRef) : "";
            if (!string.IsNullOrWhiteSpace(sh) && !string.IsNullOrWhiteSpace(sig))
                return $"chart:{sh}|series={sig}";

            return string.IsNullOrWhiteSpace(sh) ? "chart" : $"chart:{sh}";
        }

        if (spec is null)
            return new CheckResult(SectionIdForSpec(), pts, 0, false, "No chart spec provided");
        EnsureStudentZipBytes(wbS);
        var zip = _zipBytes.Value!;

        // Parse ALL charts; only pre-filter by SHEET (not name_like — that's a scored check)
        var bySheet = ParseChartsFromZip(zip);  // sheet -> charts
        IEnumerable<ChartInfo> pool = bySheet.SelectMany(kv => kv.Value);
        if (!string.IsNullOrWhiteSpace(spec.Sheet))
            pool = pool.Where(c => string.Equals(c.Sheet, spec.Sheet, StringComparison.OrdinalIgnoreCase));

        if (!pool.Any())
            return new CheckResult(SectionIdForSpec(), pts, 0, false,
                string.IsNullOrWhiteSpace(spec.Sheet)
                    ? "No charts found in workbook"
                    : $"No charts found on sheet '{spec.Sheet}'");

        // Helpers for tolerant matching
        static string NormText(string? s) =>
            System.Text.RegularExpressions.Regex.Replace((s ?? "").Trim(), @"\s+", " "); // collapse whitespace/newlines
        static string NormRef(string? a1)
        {
            var s = (a1 ?? "").Replace("$", "").ToUpperInvariant();
            int bang = s.IndexOf('!');
            return bang >= 0 ? s[(bang + 1)..] : s; // drop sheet name
        }

        ChartInfo? best = null;
        int checks = 0, hits = 0;
        double bestScore = -1;
        List<string>? bestNotes = null;

        foreach (var ch in pool)
        {
            int cks = 0, ok = 0;
            var local = new List<string>();

            void req(bool cond, string okMsg, string missMsg)
            {
                cks++;
                if (cond) { ok++; local.Add($"OK   {okMsg}"); }
                else { local.Add($"MISS {missMsg}"); }
            }

            // OPTIONAL: object name contains — now a CHECK, not a filter
            if (!string.IsNullOrWhiteSpace(spec.NameLike))
                req(ch.Name.IndexOf(spec.NameLike!, StringComparison.OrdinalIgnoreCase) >= 0,
                    $"name contains '{spec.NameLike}'",
                    $"name missing '{spec.NameLike}' (got '{ch.Name}')");

            // type
            if (!string.IsNullOrWhiteSpace(spec.Type))
                req(string.Equals(ch.Type, spec.Type, StringComparison.OrdinalIgnoreCase),
                    $"type={ch.Type}",
                    $"type got {ch.Type} expected {spec.Type}");

            // title (whitespace tolerant)
            if (!string.IsNullOrWhiteSpace(spec.Title))
                req(string.Equals(NormText(ch.Title), NormText(spec.Title), StringComparison.OrdinalIgnoreCase),
                    $"title='{ch.Title}'",
                    $"title got '{ch.Title ?? ""}' expected '{spec.Title}'");

            // title from cell (exact)
            if (!string.IsNullOrWhiteSpace(spec.TitleRef))
                req(string.Equals((ch.TitleRef ?? ""), spec.TitleRef, StringComparison.OrdinalIgnoreCase),
                    $"title_ref={ch.TitleRef}",
                    $"title_ref got {ch.TitleRef ?? ""} expected {spec.TitleRef}");

            // legend / labels / axes
            if (!string.IsNullOrWhiteSpace(spec.LegendPos))
                req(string.Equals((ch.LegendPos ?? ""), spec.LegendPos, StringComparison.OrdinalIgnoreCase),
                    $"legendPos={ch.LegendPos}",
                    $"legendPos got {ch.LegendPos ?? ""} expected {spec.LegendPos}");

            if (spec.DataLabels.HasValue)
                req(ch.DataLabels == spec.DataLabels.Value,
                    $"dataLabels={ch.DataLabels}",
                    $"dataLabels got {ch.DataLabels} expected {spec.DataLabels}");

            if (!string.IsNullOrWhiteSpace(spec.XTitle))
                req(string.Equals(NormText(ch.XTitle), NormText(spec.XTitle), StringComparison.OrdinalIgnoreCase),
                    $"xTitle='{ch.XTitle}'",
                    $"xTitle got '{ch.XTitle ?? ""}' expected '{spec.XTitle}'");

            if (!string.IsNullOrWhiteSpace(spec.YTitle))
                req(string.Equals(NormText(ch.YTitle), NormText(spec.YTitle), StringComparison.OrdinalIgnoreCase),
                    $"yTitle='{ch.YTitle}'",
                    $"yTitle got '{ch.YTitle ?? ""}' expected '{spec.YTitle}'");

            // series
            if (spec.Series is { Count: > 0 })
            {
                foreach (var exp in spec.Series)
                {
                    cks++;
                    bool found = ch.Series.Any(s =>
                        (string.IsNullOrWhiteSpace(exp.CatRef) || string.Equals(NormRef(s.CatRef), NormRef(exp.CatRef), StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrWhiteSpace(exp.ValRef) || string.Equals(NormRef(s.ValRef), NormRef(exp.ValRef), StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrWhiteSpace(exp.Name) || string.Equals(NormText(s.Name), NormText(exp.Name), StringComparison.OrdinalIgnoreCase)) &&
                        (string.IsNullOrWhiteSpace(exp.NameRef) || string.Equals(NormRef(s.NameRef), NormRef(exp.NameRef), StringComparison.OrdinalIgnoreCase))
                    );
                    if (found) { ok++; local.Add($"OK   series ({exp.CatRef} / {exp.ValRef})"); }
                    else { local.Add($"MISS series ({exp.CatRef} / {exp.ValRef})"); }
                }
            }

            // choose best by hit ratio
            double score = (cks == 0) ? 0 : (double)ok / cks;
            if (score > bestScore)
            {
                bestScore = score; best = ch; checks = cks; hits = ok; bestNotes = local;
            }
        }

        if (best is null)
            return new CheckResult(SectionIdForSpec(), pts, 0, false, "No charts found to evaluate");

        double earned = (checks == 0) ? 0 : pts * (double)hits / checks;
        bool pass = hits == checks;

        // Build concise misses list
        string missedSummary = "";
        if (!pass && bestNotes is not null)
        {
            var misses = bestNotes.Where(n => n.StartsWith("MISS"))
                                  .Select(n => n.Substring(5))
                                  .Take(6).ToList();
            if (misses.Count > 0)
            {
                int totalMiss = bestNotes.Count(n => n.StartsWith("MISS"));
                missedSummary = "; missed: " + string.Join("; ", misses) + (totalMiss > misses.Count ? " …" : "");
            }
        }

        // build the id the same way we did for early returns (so sections stay distinct)
        var successId = SectionIdForSpec();

        // build the comment (same text you already return today)
        var comment = $"matched {hits}/{checks} checks{missedSummary}; type={best.Type}; title='{best.Title ?? ""}'";

        return new CheckResult(successId, pts, earned, pass, comment);
    }

    /// <summary>
    /// Parses all charts from the student workbook OOXML zip and returns them grouped by sheet name.
    /// Extracts type, title/titleRef, legend position, data label presence, axis titles, and series refs.
    /// </summary>
    /// <param name="zipBytes">Raw student <c>.xlsx</c> bytes.</param>
    /// <returns>Dictionary of <c>sheetName → List&lt;ChartInfo&gt;</c>.</returns>
    private static Dictionary<string, List<ChartInfo>> ParseChartsFromZip(byte[] zipBytes)
    {
        using var ms = new MemoryStream(zipBytes);
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true);
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace pkg = "http://schemas.openxmlformats.org/package/2006/relationships";

        // Read an entry's text, or empty if missing.
        string ReadEntryText(string path)
        {
            var e = zip.GetEntry(path);
            if (e == null) return "";
            using var s = e.Open();
            using var r = new StreamReader(s);
            return r.ReadToEnd();
        }

        // Map sheet index → name (1-based like sheet1.xml, …)
        var wb = XDocument.Parse(ReadEntryText("xl/workbook.xml"));
        var sheetIndexToName = new Dictionary<int, string>();
        int idx = 1;
        var sheetsEl = wb.Root?.Element(nsMain + "sheets");
        if (sheetsEl != null)
            foreach (var sh in sheetsEl.Elements(nsMain + "sheet"))
                sheetIndexToName[idx++] = (string?)sh.Attribute("name") ?? $"Sheet{idx - 1}";

        var result = new Dictionary<string, List<ChartInfo>>(StringComparer.OrdinalIgnoreCase);

        for (int i = 1; i <= sheetIndexToName.Count; i++)
        {
            var sheetName = sheetIndexToName[i];
            var relsPath = $"xl/worksheets/_rels/sheet{i}.xml.rels";
            var relsTxt = ReadEntryText(relsPath);
            if (string.IsNullOrEmpty(relsTxt)) continue;

            var relsDoc = XDocument.Parse(relsTxt);
            var drawingTargets = relsDoc.Root?
                 .Elements(pkg + "Relationship")
                 .Where(r => ((string?)r.Attribute("Type"))?.EndsWith("/drawing") == true)
                 .Select(r => ((string?)r.Attribute("Target"))?.TrimStart('/').Replace("../", "xl/"))
                 .Where(t => !string.IsNullOrWhiteSpace(t))
                 .ToList() ?? new List<string>();

            foreach (var drawingRelTarget in drawingTargets)
            {
                // Normalize path like "../drawings/drawing1.xml"
                var drPath = drawingRelTarget!.StartsWith("xl/") ? drawingRelTarget! : $"xl/{drawingRelTarget}";
                var drXmlTxt = ReadEntryText(drPath);

                if (string.IsNullOrEmpty(drXmlTxt)) continue;

                var drXml = XDocument.Parse(drXmlTxt);
                // Build a map: r:id -> frame name ("Chart 1", "Chart 2", …)
                var frameMap = drXml.Descendants(xdr + "graphicFrame")
                    .Select(gf => new {
                        Name = gf.Element(xdr + "nvGraphicFramePr")?.Element(xdr + "cNvPr")?.Attribute("name")?.Value,
                        Rid = gf.Descendants(a + "graphicData").Descendants(c + "chart")
                                 .FirstOrDefault()?.Attribute(rel + "id")?.Value
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.Rid))
                    .ToDictionary(x => x.Rid!, x => x.Name ?? "Chart", StringComparer.OrdinalIgnoreCase);

                string drRelsPath = drPath.Replace("xl/drawings/", "xl/drawings/_rels/") + ".rels";
                var drRelsTxt = ReadEntryText(drRelsPath);
                if (string.IsNullOrEmpty(drRelsTxt)) continue;
                var drRels = XDocument.Parse(drRelsTxt);

                // For each <xdr:graphicFrame> → <c:chart r:id="...">
                var chartIds = drXml.Descendants(xdr + "graphicFrame")
                    .Select(gf => gf.Descendants(a + "graphicData").Descendants(c + "chart").FirstOrDefault())
                    .Where(ch => ch != null)
                    .Select(ch => (string?)ch!.Attribute(rel + "id"))
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .ToList();

                // Resolve r:id → charts/chartN.xml
                foreach (var rid in chartIds)
                {
                    var target = drRels.Root?
                        .Elements(pkg + "Relationship")
                        .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)?
                        .Attribute("Target")?.Value;

                    if (string.IsNullOrWhiteSpace(target)) continue;

                    var tgt = (target ?? "").Replace("\\", "/");
                    var chartPath = tgt.StartsWith("/") ? tgt.TrimStart('/')
                                 : tgt.StartsWith("../") ? "xl/" + tgt.Substring(3)
                                 : tgt.StartsWith("xl/") ? tgt : "xl/" + tgt;

                    var chXmlTxt = ReadEntryText(chartPath);

                    if (string.IsNullOrEmpty(chXmlTxt)) continue;

                    var chXml = XDocument.Parse(chXmlTxt);
                    var chart = new ChartInfo { Sheet = sheetName };

                    // correct name per rid
                    chart.Name = frameMap.TryGetValue(rid!, out var nm) ? nm : "Chart";

                    // detect type, then read title/legend/series...
                    var plotArea = chXml.Descendants(c + "plotArea").FirstOrDefault();
                    chart.Type = DetectChartType(plotArea);

                    // Title
                    var titleEl = chXml.Descendants(c + "title").FirstOrDefault();
                    (chart.Title, chart.TitleRef) = ReadChartText(titleEl, c, a);

                    // Axis titles (cat/val axes)
                    var catAx = plotArea?.Elements(c + "catAx").FirstOrDefault();
                    var valAx = plotArea?.Elements(c + "valAx").FirstOrDefault();
                    (chart.XTitle, _) = ReadChartText(catAx?.Element(c + "title"), c, a);
                    (chart.YTitle, _) = ReadChartText(valAx?.Element(c + "title"), c, a);

                    // Legend
                    var leg = chXml.Descendants(c + "legend").FirstOrDefault();
                    chart.LegendPos = leg?.Element(c + "legendPos")?.Attribute("val")?.Value;
                    chart.DataLabels = plotArea?.Descendants(c + "dLbls").Any() == true;

                    // Series
                    foreach (var ser in plotArea?.Descendants().Where(e => e.Name.LocalName == "ser") ?? Enumerable.Empty<XElement>())
                    {
                        var si = new SeriesInfo();
                        // name (tx)
                        var tx = ser.Element(c + "tx");
                        (si.Name, si.NameRef) = ReadChartText(tx, c, a);

                        // categories (cat)
                        var cat = ser.Element(c + "cat");
                        si.CatRef = cat?.Element(c + "strRef")?.Element(c + "f")?.Value
                                    ?? cat?.Element(c + "numRef")?.Element(c + "f")?.Value;

                        // values (val)
                        var val = ser.Element(c + "val");
                        si.ValRef = val?.Element(c + "numRef")?.Element(c + "f")?.Value
                                    ?? val?.Element(c + "strRef")?.Element(c + "f")?.Value;

                        chart.Series.Add(si);
                    }

                    if (!result.TryGetValue(sheetName, out var list)) result[sheetName] = list = new();
                    list.Add(chart);
                }
            }
        }

        return result;

        // ---- Local helpers (commented inline; XML doc tags don't apply to locals) ----

        // Reads text or cell-ref for <c:title>/<c:tx> nodes.
        static (string? txt, string? cellRef) ReadChartText(XElement? node, XNamespace cns, XNamespace ans)
        {
            if (node == null) return (null, null);

            // Accept either <c:title>/<c:tx>… or being passed <c:tx> directly
            var tx = node.Name.LocalName == "tx" ? node : node.Element(cns + "tx");
            if (tx == null) return (null, null);

            var rich = tx.Element(cns + "rich");
            if (rich != null)
            {
                var text = string.Concat(rich.Descendants(ans + "t").Select(t => t.Value));
                return (text, null);
            }

            var strRef = tx.Element(cns + "strRef");
            var f = strRef?.Element(cns + "f")?.Value;
            return (null, f);
        }

        // Infers normalized chart type from plot area.
        static string DetectChartType(XElement? plotArea)
        {
            if (plotArea == null) return "";
            XNamespace cc = "http://schemas.openxmlformats.org/drawingml/2006/chart";

            XElement? Find(string name) =>
                plotArea.Element(cc + name) ?? plotArea.Element(plotArea.GetDefaultNamespace() + name);

            // Column/Bar share barChart + barDir
            if (Find("barChart") is XElement bc)
            {
                var dirEl = bc.Element(cc + "barDir") ?? bc.Element(plotArea.GetDefaultNamespace() + "barDir");
                var dir = (string?)dirEl?.Attribute("val");
                return string.Equals(dir, "col", StringComparison.OrdinalIgnoreCase) ? "column" : "bar";
            }

            if (Find("lineChart") != null) return "line";
            if (Find("pieChart") != null) return "pie";
            if (Find("pie3DChart") != null) return "pie3D";     // 3-D pie
            if (Find("ofPieChart") != null) return "pie";        // “Pie of Pie” → count as pie
            if (Find("scatterChart") != null) return "scatter";
            if (Find("areaChart") != null) return "area";
            if (Find("doughnutChart") != null) return "doughnut";
            if (Find("radarChart") != null) return "radar";
            if (Find("bubbleChart") != null) return "bubble";
            return "";
        }
    }
}
