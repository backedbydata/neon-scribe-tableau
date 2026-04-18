namespace NeonScribe.Tableau.Core.Models;

/// <summary>
/// Represents the lineage of a field from the Tableau display name down to the data source element.
/// This captures the full path: Display Name → Internal Name → Data Source → Physical Column/Table
/// </summary>
public class FieldLineage
{
    /// <summary>
    /// The user-friendly display name shown in Tableau (e.g., "Year of Application Submit Date")
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// The internal Tableau field name (e.g., "[yr:Application Submit Date:ok]")
    /// </summary>
    public string InternalName { get; set; } = string.Empty;

    /// <summary>
    /// The base field name without derivation prefixes (e.g., "Application Submit Date")
    /// </summary>
    public string BaseFieldName { get; set; } = string.Empty;

    /// <summary>
    /// The derivation applied to the field (e.g., "Year", "Quarter", "Sum")
    /// </summary>
    public string? Derivation { get; set; }

    /// <summary>
    /// The name of the data source containing this field
    /// </summary>
    public string? DataSourceName { get; set; }

    /// <summary>
    /// The internal data source ID (e.g., "federated.1abc123xyz")
    /// </summary>
    public string? DataSourceId { get; set; }

    /// <summary>
    /// For calculated fields, the formula definition
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    /// Whether this field is a calculated field
    /// </summary>
    public bool IsCalculated { get; set; }

    /// <summary>
    /// For non-calculated fields, the physical table name in the data source
    /// </summary>
    public string? PhysicalTable { get; set; }

    /// <summary>
    /// For non-calculated fields, the physical column name in the data source
    /// </summary>
    public string? PhysicalColumn { get; set; }

    /// <summary>
    /// The data type of the field (string, integer, date, etc.)
    /// </summary>
    public string? DataType { get; set; }

    /// <summary>
    /// The role of the field (dimension or measure)
    /// </summary>
    public string? Role { get; set; }

    /// <summary>
    /// For calculated fields, the fields this calculation depends on
    /// </summary>
    public List<string>? Dependencies { get; set; }
}
