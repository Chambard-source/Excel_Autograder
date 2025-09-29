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

    // Value options
    [JsonPropertyName("case_sensitive")]
    public bool? CaseSensitive { get; set; }

    [JsonPropertyName("chart")] public ChartSpec? Chart { get; set; }

    [JsonPropertyName("table")] public TableSpec? Table { get; set; }

    [JsonPropertyName("section")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Section { get; set; }
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

    [JsonPropertyName("case_sensitive")]
    public bool? CaseSensitive { get; set; }
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

public class ChartSpec
{
    [JsonPropertyName("sheet")] public string? Sheet { get; set; }          // optional: restrict to this sheet
    [JsonPropertyName("name_like")] public string? NameLike { get; set; }   // optional: chart object name contains (e.g., "Chart 1")
    [JsonPropertyName("type")] public string? Type { get; set; }            // line, column, bar, pie, scatter, area, doughnut, radar, bubble
    [JsonPropertyName("title")] public string? Title { get; set; }          // literal title text
    [JsonPropertyName("title_ref")] public string? TitleRef { get; set; }   // e.g., "Sheet1!$B$1" (title from cell)
    [JsonPropertyName("legend_pos")] public string? LegendPos { get; set; } // t,r,b,l,tr (Excel’s legendPos)
    [JsonPropertyName("data_labels")] public bool? DataLabels { get; set; } // presence of data labels
    [JsonPropertyName("x_title")] public string? XTitle { get; set; }       // optional expected axis titles
    [JsonPropertyName("y_title")] public string? YTitle { get; set; }
    [JsonPropertyName("series")] public List<ChartSeriesSpec>? Series { get; set; }
}

public class ChartSeriesSpec
{
    [JsonPropertyName("name")] public string? Name { get; set; }            // literal
    [JsonPropertyName("name_ref")] public string? NameRef { get; set; }     // e.g., "Sheet1!$B$1"
    [JsonPropertyName("cat_ref")] public string? CatRef { get; set; }       // categories ref, e.g., "Sheet1!$A$2:$A$10"
    [JsonPropertyName("val_ref")] public string? ValRef { get; set; }       // values ref,      "Sheet1!$B$2:$B$10"
}

public class TableSpec
{
    [JsonPropertyName("sheet")] public string? Sheet { get; set; }
    [JsonPropertyName("name_like")] public string? NameLike { get; set; }
    [JsonPropertyName("columns")] public List<string>? Columns { get; set; }
    [JsonPropertyName("require_order")] public bool? RequireOrder { get; set; }

    // dimensions/range
    [JsonPropertyName("range_ref")] public string? RangeRef { get; set; }
    [JsonPropertyName("rows")] public int? Rows { get; set; }
    [JsonPropertyName("cols")] public int? Cols { get; set; }
    [JsonPropertyName("allow_extra_rows")] public bool? AllowExtraRows { get; set; }
    [JsonPropertyName("allow_extra_cols")] public bool? AllowExtraCols { get; set; }

    // content matching
    [JsonPropertyName("body_match")] public bool? BodyMatch { get; set; }
    [JsonPropertyName("body_order_matters")] public bool? BodyOrderMatters { get; set; }
    [JsonPropertyName("body_case_sensitive")] public bool? BodyCaseSensitive { get; set; }
    [JsonPropertyName("body_trim")] public bool? BodyTrim { get; set; } = true;
    [JsonPropertyName("body_rows")] public List<List<string>>? BodyRows { get; set; }

    // containment (subset) checks
    [JsonPropertyName("contains_rows")] public List<Dictionary<string, string>>? ContainsRows { get; set; }

}

public record CheckResult(string Name, double Points, double Earned, bool Passed, string Comment);
