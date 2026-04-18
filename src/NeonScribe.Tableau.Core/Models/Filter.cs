namespace NeonScribe.Tableau.Core.Models;

public class Filter
{
    public string Field { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty; // categorical, quantitative, date, parameter
    public string FilterType { get; set; } = string.Empty; // Categorical (List), Range, Relative Date, Top N
    public string ControlType { get; set; } = string.Empty; // Multi-select dropdown, Range slider, etc.
    public string? LinkedToParameter { get; set; }
    public bool? AllowMultipleValues { get; set; }
    public bool? ShowOnlyRelevantValues { get; set; }
    public List<object>? DefaultValues { get; set; }

    // Source field - the underlying data field being filtered
    public string? SourceField { get; set; }

    // Default selection - the initial/default value when known
    public string? DefaultSelection { get; set; }

    // Notes - additional context (e.g., "Dynamic - controlled by 'Additional Filter' parameter")
    public string? Notes { get; set; }

    // Scope - which visual(s) this filter applies to (null = entire dashboard)
    public string? AppliesTo { get; set; }

    // Internal param value for tracking (e.g., [federated.xxx].[none:Field:nk])
    public string? InternalParam { get; set; }

    // Field lineage - traces the field from Tableau display name to data source element
    public FieldLineage? Lineage { get; set; }

    // Position on the dashboard - used for ordering filters by visual position
    public ZonePosition? Position { get; set; }

    // For parameter-type filters: the distinct allowed values (display names in order)
    public List<string>? AllowedValues { get; set; }

    // For quantitative filters
    public object? Min { get; set; }
    public object? Max { get; set; }
    public bool? IncludeNullValues { get; set; }

    // For date filters
    public string? Period { get; set; }
    public string? Anchor { get; set; }
    public int? PeriodOffset { get; set; }

    // For Top N filters
    public int? TopN { get; set; }
    public string? By { get; set; }
    public string? Direction { get; set; } // Top or Bottom
}
