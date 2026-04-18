namespace NeonScribe.Tableau.Core.Models;

public class FieldUsage
{
    public string DataSource { get; set; } = string.Empty;
    public List<string> UsedInWorksheets { get; set; } = new();
    public List<string> UsedInFilters { get; set; } = new();
    public List<string> UsedInCalculations { get; set; } = new();
    public List<string> UsedInActions { get; set; } = new();
}

public class CalculationDependency
{
    public string Formula { get; set; } = string.Empty;
    public List<string> Dependencies { get; set; } = new();
    public List<string> UsedBy { get; set; } = new();
    public string? Explanation { get; set; }
    public string? CalculationType { get; set; }
    public string? LodType { get; set; }
}
