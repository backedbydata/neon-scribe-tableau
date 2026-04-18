namespace NeonScribe.Tableau.Core.Models;

public class TableauWorkbook
{
    public string FileName { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    public string SourceBuild { get; set; } = string.Empty;
    public int WorksheetsCount { get; set; }
    public int DashboardsCount { get; set; }
    public int StoriesCount { get; set; }
    public int DataSourcesCount { get; set; }
    public int ParametersCount { get; set; }
    public int TotalFields { get; set; }
    public int CalculatedFields { get; set; }
    public int LodCalculations { get; set; }
    public int TableCalculations { get; set; }
    public int Filters { get; set; }
    public int Actions { get; set; }

    public List<DataSource> DataSources { get; set; } = new();
    public List<Parameter> Parameters { get; set; } = new();
    public List<Worksheet> Worksheets { get; set; } = new();
    public List<Dashboard> Dashboards { get; set; } = new();

    // Cross-reference maps
    public Dictionary<string, FieldUsage> FieldUsageMap { get; set; } = new();
    public Dictionary<string, CalculationDependency> CalculationDependencies { get; set; } = new();

    // Internal lookup dictionary for name resolution
    public Dictionary<string, string> InternalNameToDisplayName { get; set; } = new();
}
