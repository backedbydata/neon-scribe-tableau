namespace NeonScribe.Tableau.Core.Models;

public class Dashboard
{
    public string Name { get; set; } = string.Empty;
    public string Caption { get; set; } = string.Empty;

    /// <summary>
    /// Title extracted from text zones (large/bold text). Stored separately for reference.
    /// </summary>
    public string? ExtractedTitle { get; set; }

    public DashboardSize Size { get; set; } = new();
    public List<DashboardZone> Zones { get; set; } = new();
    public List<DashboardAction> Actions { get; set; } = new();

    /// <summary>
    /// Color theme information for the dashboard
    /// </summary>
    public DashboardColorTheme? ColorTheme { get; set; }

    /// <summary>
    /// Dashboard-level filters (from shared-views) that apply to visuals on this dashboard
    /// </summary>
    public List<Filter> Filters { get; set; } = new();

    /// <summary>
    /// Visual components (worksheets) displayed on this dashboard (flat list for backward compatibility)
    /// </summary>
    public List<DashboardVisual> Visuals { get; set; } = new();

    /// <summary>
    /// Visual components grouped by their title text zones.
    /// Each group contains a title (from a text zone) and the worksheets that appear beneath it.
    /// </summary>
    public List<VisualGroup> VisualGroups { get; set; } = new();

    /// <summary>
    /// Supporting worksheets referenced by this dashboard (worksheets without layout-cache).
    /// Includes both shared worksheets (2+ dashboards) and background worksheets (1 dashboard).
    /// </summary>
    public List<SupportingWorksheetReference> SupportingWorksheets { get; set; } = new();
}

public class DashboardSize
{
    public int Width { get; set; }
    public int Height { get; set; }
}

public class DashboardZone
{
    public string Type { get; set; } = string.Empty; // text, worksheet, parameter-control, filter, bitmap
    public int Id { get; set; }
    public string? Content { get; set; } // For text zones
    public string? Parameter { get; set; } // For parameter-control zones
    public string? Worksheet { get; set; } // For worksheet zones
    public string? ImagePath { get; set; } // For bitmap zones
    public string? ControlType { get; set; }
    public bool? ShowApplyButton { get; set; }
    public ZonePosition Position { get; set; } = new();
    public Dictionary<string, object>? Style { get; set; }
}

public class ZonePosition
{
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
}

public class DashboardAction
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty; // filter, highlight, url, parameter
    public string Caption { get; set; } = string.Empty;
    public string Trigger { get; set; } = string.Empty; // select, hover, menu
    public bool AutoClear { get; set; }
    public ActionSource Source { get; set; } = new();
    public ActionTarget? Target { get; set; }
    public List<string>? Fields { get; set; }
    public string? UrlFormat { get; set; }
    public List<string>? FieldsUsed { get; set; }
    public string? SourceField { get; set; }
    public string? TargetParameter { get; set; }
    public string? TargetBrowser { get; set; }
}

public class ActionSource
{
    public string Type { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

public class ActionTarget
{
    public string Type { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public bool? FilterNullValues { get; set; }
}

/// <summary>
/// Reference to a supporting worksheet from a dashboard (summary info with link to details)
/// </summary>
public class SupportingWorksheetReference
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
    /// Type of visual (e.g., "Map", "Bar Chart", "KPI / Big Number")
    /// </summary>
    public string VisualType { get; set; } = string.Empty;

    /// <summary>
    /// Brief description of what this visual shows
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Whether this worksheet is shared across multiple dashboards (true) or a background worksheet (false)
    /// </summary>
    public bool IsShared { get; set; }

    /// <summary>
    /// List of dashboard names where this worksheet is used (populated only for shared worksheets)
    /// </summary>
    public List<string> UsedInDashboards { get; set; } = new();
}

/// <summary>
/// Color theme information for dashboards
/// </summary>
public class DashboardColorTheme
{
    /// <summary>
    /// Primary/base color for the dashboard (from format color attribute)
    /// </summary>
    public string? PrimaryColor { get; set; }

    /// <summary>
    /// List of distinct mark colors used across all visuals on this dashboard
    /// </summary>
    public List<string> MarkColors { get; set; } = new();

    /// <summary>
    /// Named palette used (if any)
    /// </summary>
    public string? PaletteName { get; set; }
}
