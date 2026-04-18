namespace NeonScribe.Tableau.Core.Models;

public class Parameter
{
    public string Name { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string Caption { get; set; } = string.Empty;
    public string DataType { get; set; } = string.Empty;
    public string DomainType { get; set; } = string.Empty; // range, list, all
    public object? CurrentValue { get; set; }
    public object? DefaultValue { get; set; }
    public ParameterAllowableValues? AllowableValues { get; set; }

    // Display name for the default value (resolved from alias)
    public string? DefaultDisplayName { get; set; }

    // Maps value -> display alias (e.g., 1 -> "Age Group", 2 -> "Prior Edu")
    public Dictionary<string, string>? ValueAliases { get; set; }
}

public class ParameterAllowableValues
{
    public string Type { get; set; } = string.Empty; // range or list

    // For range type
    public object? Min { get; set; }
    public object? Max { get; set; }
    public object? Granularity { get; set; }

    // For list type
    public List<object>? Values { get; set; }

    // For list type - display names for values
    public List<string>? DisplayNames { get; set; }
}
