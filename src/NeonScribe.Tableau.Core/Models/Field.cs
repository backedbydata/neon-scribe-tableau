namespace NeonScribe.Tableau.Core.Models;

public class Field
{
    public string Name { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string DataType { get; set; } = string.Empty;
    public string Role { get; set; } = string.Empty; // dimension or measure
    public string Type { get; set; } = string.Empty; // nominal, ordinal, quantitative
    public bool IsCalculated { get; set; }
    public string? Formula { get; set; }
    public List<string>? Dependencies { get; set; }
    public string? CalculationType { get; set; } // LOD, Table Calculation, etc.
    public string? LodType { get; set; } // FIXED, INCLUDE, EXCLUDE
    public List<string>? LodDimensions { get; set; } // Dimensions in LOD scope
    public string? Explanation { get; set; } // Natural language explanation
    public string? TableCalcFunction { get; set; } // Table calc function name (RUNNING_SUM, WINDOW_AVG, etc.)
}
