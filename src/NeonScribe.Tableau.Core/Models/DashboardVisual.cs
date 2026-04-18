namespace NeonScribe.Tableau.Core.Models;

/// <summary>
/// Represents a visual component (worksheet) displayed on a dashboard
/// </summary>
public class DashboardVisual
{
    /// <summary>
    /// Internal name of the worksheet
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// User-friendly caption/display name
    /// </summary>
    public string Caption { get; set; } = string.Empty;

    /// <summary>
    /// Type of visual (e.g., "KPI Metric", "Bar Chart", "Line Chart", "Heat Map", etc.)
    /// </summary>
    public string VisualType { get; set; } = string.Empty;

    /// <summary>
    /// Reference to the actual worksheet object
    /// </summary>
    public string WorksheetReference { get; set; } = string.Empty;

    /// <summary>
    /// List of filter field names that apply to this visual
    /// </summary>
    public List<string> UsesFilters { get; set; } = new();

    /// <summary>
    /// Fields used in this visual (dimensions and measures) - simple list for backward compatibility
    /// </summary>
    public List<string> FieldsUsed { get; set; } = new();

    /// <summary>
    /// Detailed field usage with shelf and aggregation information
    /// </summary>
    public List<FieldUsageInSheet> DetailedFieldsUsed { get; set; } = new();

    /// <summary>
    /// Custom tooltip configuration for this visual
    /// </summary>
    public Tooltip? Tooltip { get; set; }

    /// <summary>
    /// Mark type used in this visual (e.g., "Bar", "Line", "Area", "Pie", "Automatic")
    /// </summary>
    public string MarkType { get; set; } = string.Empty;

    /// <summary>
    /// Mark encodings (Color, Size, Text, Shape, Detail, Tooltip fields)
    /// </summary>
    public MarkEncodings? MarkEncodings { get; set; }

    /// <summary>
    /// Whether this worksheet appears on multiple dashboards
    /// </summary>
    public bool IsSharedAcrossDashboards { get; set; }

    /// <summary>
    /// Position and size of the visual on the dashboard
    /// </summary>
    public ZonePosition Position { get; set; } = new();

    /// <summary>
    /// Brief description of what this visual shows (can be auto-generated or user-provided)
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Filters that apply only to this specific visual (not dashboard-wide)
    /// </summary>
    public List<Filter>? VisualSpecificFilters { get; set; }

    /// <summary>
    /// Customized label structure for KPI displays
    /// </summary>
    public CustomizedLabel? CustomizedLabel { get; set; }

    /// <summary>
    /// Whether this visual has actual tooltip content (not just mark encoding fields)
    /// </summary>
    public bool HasActualTooltipContent { get; set; }

    /// <summary>
    /// Map configuration settings (for geographic visualizations)
    /// </summary>
    public MapConfiguration? MapConfiguration { get; set; }

    /// <summary>
    /// Title from layout-options (if different from caption/name)
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Table configuration settings (for table/crosstab visualizations)
    /// </summary>
    public TableConfiguration? TableConfiguration { get; set; }
}
