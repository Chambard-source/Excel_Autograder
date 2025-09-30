using ClosedXML.Excel;
using System.Globalization;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(args);

// Allow large uploads (50 MB here, adjust if needed)
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = 500 * 1024 * 1024);
var app = builder.Build();
Console.WriteLine($"ContentRootPath = {app.Environment.ContentRootPath}");
Console.WriteLine($"WebRootPath     = {app.Environment.WebRootPath}");

app.UseDefaultFiles();   // serves wwwroot/index.html
app.UseStaticFiles();

// POST /api/grade
app.MapPost("/api/grade", async (HttpRequest req) =>
{
    if (!req.HasFormContentType)
        return Results.BadRequest(new { error = "Expected multipart/form-data" });

    var form = await req.ReadFormAsync();

    // Required files
    var key = form.Files.GetFile("key");
    var students = form.Files.GetFiles("students");

    if (key is null || students.Count == 0)
        return Results.BadRequest(new { error = "Upload key.xlsx and at least one student workbook." });

    // --- Get rubric JSON ---
    // 1) Try file part "rubricJson" (what the UI sends from the preview box)
    // 2) Fallback to file part "rubric" (legacy)
    // 3) Fallback to text field "rubric_json"
    string? rubricText = null;

    var rubricJsonFile = form.Files.GetFile("rubricJson")
                        ?? form.Files.GetFile("rubric");

    if (rubricJsonFile != null)
    {
        using var rr = new StreamReader(rubricJsonFile.OpenReadStream());
        rubricText = await rr.ReadToEndAsync();
    }
    else
    {
        rubricText = form["rubric_json"].FirstOrDefault();
    }

    if (string.IsNullOrWhiteSpace(rubricText))
        return Results.BadRequest(new { error = "Provide a rubric: upload as 'rubricJson' (or 'rubric') or send text as 'rubric_json'." });

    Rubric? rub;
    try
    {
        rub = JsonSerializer.Deserialize<Rubric>(
            rubricText!,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true }
        );
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = "Invalid rubric JSON", detail = ex.Message });
    }
    if (rub is null)
        return Results.BadRequest(new { error = "Rubric could not be parsed." });

    // Open key workbook
    using var keyStream = key.OpenReadStream();
    using var wbKey = new XLWorkbook(keyStream);

    // Grade students (parallel with throttling)
    var results = new List<object>();
    var throttler = new SemaphoreSlim(4); // tune: 4-8
    var tasks = new List<Task>();

    // Grade students (parallel with throttling)
    foreach (var s in students)
    {
        await throttler.WaitAsync();
        tasks.Add(Task.Run(async () =>
        {
            try
            {
                // --- BUFFER STUDENT (so Grader can inspect ZIP for CFs)
                byte[] sBytes;
                using (var ms = new MemoryStream())
                {
                    using var sStreamRaw = s.OpenReadStream();
                    await sStreamRaw.CopyToAsync(ms);
                    sBytes = ms.ToArray();
                }

                using var wbStudent = new XLWorkbook(new MemoryStream(sBytes));
                var grade = Grader.Run(wbKey, wbStudent, rub, sBytes); // <-- new overload
                lock (results) results.Add(new { student = s.FileName, grade });
            }
            catch (Exception ex)
            {
                lock (results) results.Add(new { student = s.FileName, error = ex.Message });
            }
            finally { throttler.Release(); }
        }));
    }

    await Task.WhenAll(tasks);
    return Results.Json(results, new JsonSerializerOptions { WriteIndented = true });
});

// --- Rubric persistence (save files under wwwroot/rubrics) ---
var rubricDir = Path.Combine(app.Environment.WebRootPath ?? "wwwroot", "rubrics");
Directory.CreateDirectory(rubricDir);

app.MapGet("/api/rubrics", () =>
{
    var files = Directory.GetFiles(rubricDir, "*.json")
        .Select(p => Path.GetFileName(p));
    return Results.Ok(files);
});

string SanitizeFileName(string name)
{
    var onlyName = Path.GetFileName(name);
    foreach (var c in Path.GetInvalidFileNameChars())
        onlyName = onlyName.Replace(c, '_');
    if (string.IsNullOrWhiteSpace(onlyName)) onlyName = "rubric.json";
    if (!onlyName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
        onlyName += ".json";
    return onlyName;
}

app.MapGet("/api/rubric/{name}", (string name) =>
{
    var safe = SanitizeFileName(name);
    var path = Path.Combine(rubricDir, safe);
    if (!System.IO.File.Exists(path)) return Results.NotFound();
    var json = System.IO.File.ReadAllText(path);
    return Results.Text(json, "application/json");
});

app.MapPost("/api/rubric/{name}", async (string name, HttpRequest req) =>
{
    using var reader = new StreamReader(req.Body);
    var json = await reader.ReadToEndAsync();

    try
    {
        _ = JsonSerializer.Deserialize<Rubric>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = "Invalid rubric JSON", detail = ex.Message });
    }

    var safe = SanitizeFileName(name);
    var path = Path.Combine(rubricDir, safe);
    await System.IO.File.WriteAllTextAsync(path, json);
    return Results.Ok(new { saved = safe });
});


// POST /api/rubric/auto  (multipart/form-data: key=<file>, sheet=<optional>, all=<optional "true">, total=<optional>)
app.MapPost("/api/rubric/auto", async (HttpRequest req) =>
{
    try
    {
        if (!req.HasFormContentType) return Results.BadRequest(new { error = "Expected multipart/form-data" });
        var form = await req.ReadFormAsync();
        var keyFile = form.Files.GetFile("key");
        if (keyFile is null) return Results.BadRequest(new { error = "Upload key workbook as 'key'" });

        string? sheet = form["sheet"].FirstOrDefault() ?? req.Query["sheet"].FirstOrDefault() ?? req.Query["sheetHint"].FirstOrDefault();
        string? allStr = form["all"].FirstOrDefault() ?? req.Query["all"].FirstOrDefault() ?? req.Query["allSheets"].FirstOrDefault();
        string? totalStr = form["total"].FirstOrDefault() ?? req.Query["total"].FirstOrDefault();

        static bool ParseBool(string? s) =>
            !string.IsNullOrWhiteSpace(s) &&
            (s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
             s.Equals("on", StringComparison.OrdinalIgnoreCase) ||
             s == "1");

        bool allSheets = ParseBool(allStr);

        // ---- buffer the key file into memory
        byte[] keyBytes;
        using (var ms = new MemoryStream())
        {
            using var src = keyFile.OpenReadStream();
            await src.CopyToAsync(ms);
            keyBytes = ms.ToArray();
        }

        // build from key
        using var wbKey = new XLWorkbook(new MemoryStream(keyBytes));
        var rub = RubricAuto.BuildFromKey(wbKey, sheet, allSheets, 0, keyBytes);

        // optional scaling
        if (double.TryParse(totalStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var desired) && desired > 0)
            RubricAuto.ScalePoints(rub, desired);

        return Results.Json(rub, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine(ex); // log to server console
        return Results.Problem(
            title: "Auto-rubric generation failed",
            detail: ex.ToString(),
            statusCode: 500
        );
    }
});

// POST /api/auto-rubric  (same behavior; some UIs call this route)
app.MapPost("/api/auto-rubric", async (HttpRequest req) =>
{
    try
    {
        if (!req.HasFormContentType) return Results.BadRequest(new { error = "Expected multipart/form-data" });
        var form = await req.ReadFormAsync();
        var keyFile = form.Files.GetFile("key");
        if (keyFile is null) return Results.BadRequest(new { error = "Upload key workbook as 'key'" });

        string? sheet = form["sheet"].FirstOrDefault() ?? req.Query["sheet"].FirstOrDefault() ?? req.Query["sheetHint"].FirstOrDefault();
        string? allStr = form["all"].FirstOrDefault() ?? req.Query["all"].FirstOrDefault() ?? req.Query["allSheets"].FirstOrDefault();
        string? totalStr = form["total"].FirstOrDefault() ?? req.Query["total"].FirstOrDefault();

        static bool ParseBool(string? s) =>
            !string.IsNullOrWhiteSpace(s) &&
            (s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
             s.Equals("on", StringComparison.OrdinalIgnoreCase) ||
             s == "1");

        bool allSheets = ParseBool(allStr);

        // ---- buffer the key file into memory
        byte[] keyBytes;
        using (var ms = new MemoryStream())
        {
            using var src = keyFile.OpenReadStream();
            await src.CopyToAsync(ms);
            keyBytes = ms.ToArray();
        }

        // build from key
        using var wbKey = new XLWorkbook(new MemoryStream(keyBytes));
        var rub = RubricAuto.BuildFromKey(wbKey, sheet, allSheets, 0, keyBytes);

        // optional scaling
        if (double.TryParse(totalStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var desired) && desired > 0)
            RubricAuto.ScalePoints(rub, desired);

        return Results.Json(rub, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine(ex);
        return Results.Problem(
            title: "Auto-rubric generation failed",
            detail: ex.ToString(),
            statusCode: 500
        );
    }
});

app.MapPost("/api/key/sheets", async (HttpRequest req) =>
{
    if (!req.HasFormContentType) return Results.BadRequest(new { error = "Expected multipart/form-data" });
    var form = await req.ReadFormAsync();
    var keyFile = form.Files.GetFile("key");
    if (keyFile is null) return Results.BadRequest(new { error = "Upload key workbook as 'key'" });

    using var ms = new MemoryStream();
    using (var src = keyFile.OpenReadStream()) { await src.CopyToAsync(ms); }
    var keyBytes = ms.ToArray();

    using var wbKey = new XLWorkbook(new MemoryStream(keyBytes));
    var names = wbKey.Worksheets.Select(w => w.Name).ToList();
    return Results.Json(names);
});

// Build rubric from either flat ranges or named sections + ranges
app.MapPost("/api/rubric/from-ranges", async (HttpRequest req) =>
{
    if (!req.HasFormContentType)
        return Results.BadRequest(new { error = "Expected multipart/form-data" });

    var form = await req.ReadFormAsync();
    var key = form.Files.GetFile("key");
    if (key is null)
        return Results.BadRequest(new { error = "Missing 'key' file" });

    // read workbook bytes once (also used for ZIP artifact scan)
    byte[] keyBytes;
    using (var ms = new MemoryStream())
    {
        await key.CopyToAsync(ms);
        keyBytes = ms.ToArray();
    }

    using var wbKey = new ClosedXML.Excel.XLWorkbook(new MemoryStream(keyBytes));

    var includeArtifacts = string.Equals(form["include_artifacts"], "true", StringComparison.OrdinalIgnoreCase);
    double.TryParse(form["total"], out var targetTotal);

    // Accept either ranges_json or sections_json
    var rangesJson = form["ranges_json"].ToString();
    var sectionsJson = form["sections_json"].ToString();

    try
    {
        if (!string.IsNullOrWhiteSpace(sectionsJson))
        {
            // { "SheetA": [ { "name": "Drink Types", "ranges": ["D8:D17"] }, ... ] }
            var opts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };

            var dto = System.Text.Json.JsonSerializer.Deserialize<
                Dictionary<string, List<SectionDto>>
            >(sectionsJson, opts) ?? new();

            var sectionsPerSheet =
                                dto.ToDictionary(
                                    kv => kv.Key,
                                    kv => kv.Value
                                          .Select(s => (section: (s.Name ?? s.Section ?? "Section"), ranges: (s.Ranges ?? new List<string>())))
                                          .ToList(),
                                    StringComparer.OrdinalIgnoreCase);

            var rub = RubricAuto.BuildFromKeyRanges(wbKey, sectionsPerSheet, includeArtifacts, targetTotal, keyBytes);
            return Results.Json(rub);
        }
        else if (!string.IsNullOrWhiteSpace(rangesJson))
        {
            // { "SheetA": ["A2:B20","E8", ...] }
            var opts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var ranges = System.Text.Json.JsonSerializer.Deserialize<
                Dictionary<string, List<string>>
            >(rangesJson, opts) ?? new();

            var rub = RubricAuto.BuildFromKeyRanges(wbKey, ranges, includeArtifacts, targetTotal, keyBytes);
            return Results.Json(rub);
        }
        else
        {
            return Results.BadRequest(new { error = "Provide either ranges_json or sections_json" });
        }
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = ex.Message });
    }
});


app.Run();


// at bottom of Program.cs (after app.Run())
public sealed class SectionDto
{
    public string? Name { get; set; }
    public string? Section { get; set; }  // allow "section" from JSON too
    public List<string>? Ranges { get; set; }
    public string ResolvedName => string.IsNullOrWhiteSpace(Name) ? (Section ?? "Section") : Name;
}
