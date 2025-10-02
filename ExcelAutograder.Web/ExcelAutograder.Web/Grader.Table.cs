using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

public static partial class Grader
{
    private static CheckResult GradeTable(Rule rule, IXLWorksheet wsS) 
    {
        var pts = rule.Points;
        var spec = rule.Table;
        if (spec is null)
            return new CheckResult("table", pts, 0, false, "No table spec provided");

        // Helper to build a section-aware id like "table:Sales/MyTableLike"
        string TableId() =>
            (!string.IsNullOrWhiteSpace(spec.Sheet) || !string.IsNullOrWhiteSpace(spec.NameLike))
            ? $"table:{spec.Sheet ?? ""}{(!string.IsNullOrWhiteSpace(spec.Sheet) && !string.IsNullOrWhiteSpace(spec.NameLike) ? "/" : "")}{spec.NameLike ?? ""}"
            : "table";

        // Sheet gating
        if (!string.IsNullOrWhiteSpace(spec.Sheet) &&
            !string.Equals(spec.Sheet, wsS.Name, StringComparison.OrdinalIgnoreCase))
        {
            return new CheckResult(TableId(), pts, 0, false,
                $"Expected on sheet '{spec.Sheet}', grading '{wsS.Name}'");
        }

        // Candidate tables
        var tables = wsS.Tables.AsEnumerable();
        if (!string.IsNullOrWhiteSpace(spec.NameLike))
            tables = tables.Where(t => t.Name.IndexOf(spec.NameLike!, StringComparison.OrdinalIgnoreCase) >= 0);

        if (!tables.Any())
            return new CheckResult(TableId(), pts, 0, false,
                $"No table {(string.IsNullOrWhiteSpace(spec.NameLike) ? "" : $"matching '{spec.NameLike}' ")}found on '{wsS.Name}'");

        int checksTotal = 0, bestHits = -1;
        string bestNote = "", bestName = "";

        foreach (var t in tables)
        {
            var notes = new List<string>();
            int cks = 0, hits = 0;

            // ---------- HEADERS ----------
            var headers = t.Fields.Select(f => f.Name?.Trim() ?? "").ToList();
            if (spec.Columns is { Count: > 0 })
            {
                foreach (var want in spec.Columns)
                {
                    cks++;
                    bool present = headers.Any(h => string.Equals(h, want, StringComparison.OrdinalIgnoreCase));
                    if (present) { hits++; notes.Add($"[{t.Name}] has '{want}'"); }
                    else notes.Add($"[{t.Name}] missing '{want}'");
                }
                if (spec.RequireOrder == true)
                {
                    cks++;
                    bool orderOk = true; int last = -1;
                    foreach (var want in spec.Columns)
                    {
                        int idx = headers.FindIndex(h => string.Equals(h, want, StringComparison.OrdinalIgnoreCase));
                        if (idx < 0 || idx < last) { orderOk = false; break; }
                        last = idx;
                    }
                    if (orderOk) { hits++; notes.Add($"[{t.Name}] column order ok"); }
                    else notes.Add($"[{t.Name}] column order wrong");
                }
            }

            // ---------- RANGE REF ----------
            if (!string.IsNullOrWhiteSpace(spec.RangeRef))
            {
                cks++;
                bool ok = false;
                try
                {
                    string sheetPart = wsS.Name, addrPart = spec.RangeRef!;
                    int bang = spec.RangeRef!.IndexOf('!');
                    if (bang >= 0) { sheetPart = spec.RangeRef!.Substring(0, bang); addrPart = spec.RangeRef!.Substring(bang + 1); }
                    var sh = string.Equals(sheetPart, wsS.Name, StringComparison.OrdinalIgnoreCase)
                        ? wsS : wsS.Workbook.Worksheet(sheetPart);

                    var expected = sh.Range(addrPart).RangeAddress;
                    var got = t.RangeAddress;
                    ok = got.FirstAddress.RowNumber == expected.FirstAddress.RowNumber
                      && got.FirstAddress.ColumnNumber == expected.FirstAddress.ColumnNumber
                      && got.LastAddress.RowNumber == expected.LastAddress.RowNumber
                      && got.LastAddress.ColumnNumber == expected.LastAddress.ColumnNumber;
                }
                catch { ok = false; }

                if (ok) { hits++; notes.Add($"[{t.Name}] range matches {spec.RangeRef}"); }
                else notes.Add($"[{t.Name}] range != {spec.RangeRef} (got {t.RangeAddress.ToStringRelative()})");
            }

            // ---------- DIMENSIONS (data body only) ----------
            var body = t.DataRange;
            int bodyRows = body?.RowCount() ?? 0;
            int bodyCols = body?.ColumnCount() ?? 0;

            if (spec.Rows.HasValue)
            {
                cks++;
                bool ok = (spec.AllowExtraRows == true) ? (bodyRows >= spec.Rows.Value) : (bodyRows == spec.Rows.Value);
                if (ok) { hits++; notes.Add($"rows {bodyRows} ok"); }
                else notes.Add($"rows {bodyRows} not {(spec.AllowExtraRows == true ? ">=" : "=")} {spec.Rows}");
            }
            if (spec.Cols.HasValue)
            {
                cks++;
                bool ok = (spec.AllowExtraCols == true) ? (bodyCols >= spec.Cols.Value) : (bodyCols == spec.Cols.Value);
                if (ok) { hits++; notes.Add($"cols {bodyCols} ok"); }
                else notes.Add($"cols {bodyCols} not {(spec.AllowExtraCols == true ? ">=" : "=")} {spec.Cols}");
            }

            // ---------- CONTAINS ROWS ----------
            if (spec.ContainsRows is { Count: > 0 })
            {
                var idxByName = headers.Select((h, i) => (h, i))
                    .ToDictionary(x => x.h, x => x.i, StringComparer.OrdinalIgnoreCase);

                foreach (var required in spec.ContainsRows)
                {
                    cks++;
                    bool found = false;
                    if (body != null)
                    {
                        foreach (var row in body.Rows())
                        {
                            bool match = true;
                            foreach (var kv in required)
                            {
                                if (!idxByName.TryGetValue(kv.Key, out int ci)) { match = false; break; }
                                var text = row.Cell(ci + 1).GetFormattedString()?.Trim() ?? "";
                                if (!string.Equals(text, (kv.Value ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                                { match = false; break; }
                            }
                            if (match) { found = true; break; }
                        }
                    }
                    if (found) { hits++; notes.Add($"contains: {string.Join(", ", required.Select(kv => $"{kv.Key}='{kv.Value}'"))}"); }
                    else notes.Add($"missing: {string.Join(", ", required.Select(kv => $"{kv.Key}='{kv.Value}'"))}");
                }
            }

            // ---------- WHOLE-BODY COMPARISON ----------
            if (spec.BodyMatch == true && spec.BodyRows is { Count: > 0 })
            {
                cks++;

                var sBody = new List<List<string>>();
                if (body != null)
                {
                    foreach (var r in body.Rows())
                    {
                        var rowVals = new List<string>();
                        foreach (var c in r.Cells()) rowVals.Add(c.GetFormattedString() ?? "");
                        sBody.Add(rowVals);
                    }
                }

                bool trim = spec.BodyTrim != false;
                bool caseSens = spec.BodyCaseSensitive == true;
                string Norm(string x) => trim ? (x ?? "").Trim() : (x ?? "");
                StringComparer cmp = caseSens ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

                bool shapeOk = sBody.Count == spec.BodyRows.Count &&
                               (sBody.Count == 0 || sBody[0].Count == spec.BodyRows[0].Count);

                bool match = shapeOk;
                if (match)
                {
                    if (spec.BodyOrderMatters == true)
                    {
                        for (int i = 0; i < sBody.Count && match; i++)
                            for (int j = 0; j < sBody[i].Count && match; j++)
                                if (!cmp.Equals(Norm(sBody[i][j]), Norm(spec.BodyRows[i][j])))
                                    match = false;
                    }
                    else
                    {
                        string Key(List<string> row) => string.Join("\u001F", row.Select(Norm));
                        var left = sBody.Select(Key).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
                        var right = spec.BodyRows.Select(Key).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
                        match = left.Count == right.Count && left.All(kv => right.TryGetValue(kv.Key, out int n) && n == kv.Value);
                    }
                }

                if (match) { hits++; notes.Add("body matches"); }
                else notes.Add("body does not match");
            }

            // Best candidate scoring
            if (cks > 0 && hits > bestHits)
            {
                bestHits = hits;
                checksTotal = cks;
                bestName = t.Name;
                bestNote = string.Join(" | ", notes);
            }
        }

        if (checksTotal == 0)
            return new CheckResult(TableId(), pts, 0, false,
                "No checks declared (add columns / range_ref / rows/cols / contains_rows / body_match).");

        double frac = (double)bestHits / checksTotal;
        double earned = pts * frac;
        bool pass = Math.Abs(frac - 1.0) < 1e-9;

        // Success id includes sheet + nameLike (so it groups under your section)
        return new CheckResult($"table:{wsS.Name}/{(spec.NameLike ?? bestName)}", pts, earned, pass, bestNote);
    }
}
