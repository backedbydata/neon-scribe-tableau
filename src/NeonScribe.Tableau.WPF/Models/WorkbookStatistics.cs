namespace NeonScribe.Tableau.WPF.Models;

public class WorkbookStatistics
{
    public string FileName { get; set; } = string.Empty;
    public string WorkbookName { get; set; } = string.Empty;
    public int DataSourceCount { get; set; }
    public int WorksheetCount { get; set; }
    public int DashboardCount { get; set; }
    public int ParameterCount { get; set; }
    public int TotalFieldCount { get; set; }
}
