using System.Text.Json;
using System.Text.Json.Serialization;

public class Rubric
{
    [JsonPropertyName("points")] public double Points { get; set; } = 100;
    [JsonPropertyName("sheets")] public Dictionary<string, SheetSpec> Sheets { get; set; } = new();
    [JsonPropertyName("scoring")] public Scoring? Scoring { get; set; }
    [JsonPropertyName("report")] public Report? Report { get; set; }
}

public class SheetSpec { [JsonPropertyName("checks")] public List<Rule> Checks { get; set; } = new(); }
public class Scoring { [JsonPropertyName("round_to")] public double? RoundTo { get; set; } }
public class Report { [JsonPropertyName("include_pass_fail_column")] public bool IncludePassFailColumn { get; set; } = true; [JsonPropertyName("include_comments")] public bool IncludeComments { get; set; } = true; }

public class Rule
{
    // Common
    [JsonPropertyName("type")] public string Type { get; set; } = "";
    [JsonPropertyName("points")] public double Points { get; set; }
    [JsonPropertyName("cell")] public string? Cell { get; set; }
    [JsonPropertyName("range")] public string? Range { get; set; }
    [JsonPropertyName("tolerance")] public double? Tolerance { get; set; }
    [JsonPropertyName("any_of")] public List<RuleOption>? AnyOf { get; set; }
    [JsonPropertyName("expected_from_key")] public bool? ExpectedFromKey { get; set; }
    [JsonPropertyName("note")] public string? Note { get; set; }

    // Value
    [JsonPropertyName("expected")] public JsonElement? Expected { get; set; }
    [JsonPropertyName("expected_regex")] public string? ExpectedRegex { get; set; }

    // Formula
    [JsonPropertyName("expected_formula")] public string? ExpectedFormula { get; set; }
    [JsonPropertyName("allow_regex")] public bool? AllowRegex { get; set; }
    [JsonPropertyName("expected_formula_regex")] public string? ExpectedFormulaRegex { get; set; }

    // Format
    [JsonPropertyName("format")] public FormatSpec? Format { get; set; }

    // Custom
    [JsonPropertyName("require")] public RequireSpec? Require { get; set; }

    // Sequence helpers (for range_sequence)
    [JsonPropertyName("start")] public double? Start { get; set; }
    [JsonPropertyName("step")] public double? Step { get; set; }

    // Formula style requirement
    [JsonPropertyName("require_absolute")] public bool? RequireAbsolute { get; set; }

    // Pivot layout (for rule type "pivot_layout")
    [JsonPropertyName("pivot")] public PivotSpec? Pivot { get; set; }

    [JsonPropertyName("cond")] public ConditionalFormatSpec? Cond { get; set; }

}


public class RuleOption
{
    // for value checks
    [JsonPropertyName("expected")] public JsonElement? Expected { get; set; }
    [JsonPropertyName("expected_regex")] public string? ExpectedRegex { get; set; }

    // for format checks
    [JsonPropertyName("format")] public FormatSpec? Format { get; set; }

    // NEW: for formula checks
    [JsonPropertyName("expected_formula")] public string? ExpectedFormula { get; set; }
    [JsonPropertyName("expected_formula_regex")] public string? ExpectedFormulaRegex { get; set; }
    [JsonPropertyName("expected_from_key")] public bool? ExpectedFromKey { get; set; }
    [JsonPropertyName("require_absolute")] public bool? RequireAbsolute { get; set; }
}


public class FormatSpec
{
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? NumberFormat { get; set; }
    public FillSpec? Fill { get; set; }
    public FontSpec? Font { get; set; }
    public AlignmentSpec? Alignment { get; set; }
    public BorderSpec? Border { get; set; }
}

public class FillSpec { public string? Rgb { get; set; } }
public class FontSpec { public string? Name { get; set; } public double? Size { get; set; } public bool? Bold { get; set; } public bool? Italic { get; set; } }
public class AlignmentSpec { public string? Horizontal { get; set; } public string? Vertical { get; set; } }
public class BorderSpec { public bool? Outline { get; set; } }
public class RequireSpec { public string? Sheet { get; set; } public string? PivotTableLike { get; set; } }
public class PivotSpec
{
    [JsonPropertyName("sheet")] public string? Sheet { get; set; }
    [JsonPropertyName("tableNameLike")] public string? TableNameLike { get; set; }
    [JsonPropertyName("rows")] public List<string>? Rows { get; set; }
    [JsonPropertyName("columns")] public List<string>? Columns { get; set; }
    [JsonPropertyName("filters")] public List<string>? Filters { get; set; }
    [JsonPropertyName("values")] public List<PivotValueSpec>? Values { get; set; }
}

public class PivotValueSpec
{
    [JsonPropertyName("field")] public string Field { get; set; } = "";
    [JsonPropertyName("agg")] public string? Agg { get; set; } // sum,count,average,min,max
}

public class ConditionalFormatSpec
{
    [JsonPropertyName("sheet")] public string? Sheet { get; set; }        // optional; if null, we’ll match on the sheet in the rubric section
    [JsonPropertyName("range")] public string? Range { get; set; }        // optional; e.g., "B2:B50" (we’ll consider overlap)
    [JsonPropertyName("type")] public string? Type { get; set; }          // cellIs, expression, containsText, top10, dataBar, colorScale, iconSet
    [JsonPropertyName("op")] public string? Operator { get; set; }        // gt, ge, lt, le, eq, ne, between, notBetween (for cellIs)
    [JsonPropertyName("formula1")] public string? Formula1 { get; set; }  // first formula (as text, e.g. "=B2>0")
    [JsonPropertyName("formula2")] public string? Formula2 { get; set; }  // second formula for between/notBetween
    [JsonPropertyName("text")] public string? Text { get; set; }          // used by containsText rules
    [JsonPropertyName("fillRgb")] public string? FillRgb { get; set; }    // optional (e.g. "FFFF00"); we’ll try to match
}


public record CheckResult(string Name, double Points, double Earned, bool Passed, string Comment);
