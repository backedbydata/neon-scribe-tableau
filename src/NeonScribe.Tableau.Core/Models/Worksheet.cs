namespace NeonScribe.Tableau.Core.Models;

public class Worksheet
{
    public string Name { get; set; } = string.Empty;
    public string Caption { get; set; } = string.Empty;
    public string VisualType { get; set; } = string.Empty;
    public string MarkType { get; set; } = string.Empty;
    public List<FieldUsageInSheet> FieldsUsed { get; set; } = new();
    public List<Filter> Filters { get; set; } = new();
    public Tooltip? Tooltip { get; set; }

    /// <summary>
    /// Mark encoding shelf information (Color, Size, Text, Detail, Shape)
    /// </summary>
    public MarkEncodings? MarkEncodings { get; set; }

    /// <summary>
    /// Whether this worksheet is hidden in Tableau (from window element's hidden attribute)
    /// </summary>
    public bool IsHidden { get; set; }

    /// <summary>
    /// Customized label structure for KPI displays (parsed from customized-label element)
    /// </summary>
    public CustomizedLabel? CustomizedLabel { get; set; }

    /// <summary>
    /// Whether this worksheet has actual tooltip content (not just mark encoding fields)
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

public class FieldUsageInSheet
{
    public string Field { get; set; } = string.Empty;
    public string Shelf { get; set; } = string.Empty; // Rows, Columns, Color, Size, etc.
    public string Aggregation { get; set; } = string.Empty; // SUM, AVG, COUNT, None, etc.
}

public class Tooltip
{
    public bool HasCustomTooltip { get; set; }
    public string Content { get; set; } = string.Empty;
    public List<string> FieldsUsed { get; set; } = new();
}

/// <summary>
/// Contains all mark encodings for a worksheet
/// </summary>
public class MarkEncodings
{
    public ColorEncoding? Color { get; set; }
    public SizeEncoding? Size { get; set; }
    public TextEncoding? Text { get; set; }
    public ShapeEncoding? Shape { get; set; }
    public List<string> DetailFields { get; set; } = new();
    public List<string> TooltipFields { get; set; } = new();
}

public class ColorEncoding
{
    public string Field { get; set; } = string.Empty;
    /// <summary>Raw internal field reference, e.g., [none:Calculation_789818813289242634:nk]</summary>
    public string? RawField { get; set; }
    public string PaletteType { get; set; } = string.Empty; // palette, custom-interpolated
    public string? PaletteName { get; set; }
    /// <summary>
    /// Color mappings: value -> hex color
    /// </summary>
    public Dictionary<string, string> ColorMappings { get; set; } = new();
    /// <summary>
    /// Mark color (the primary visualization color, e.g., #820000)
    /// </summary>
    public string? MarkColor { get; set; }
    /// <summary>
    /// Optional grouped color mappings for CASE-based dynamic fields.
    /// Key = parameter value label (e.g., "Race", "Sex"), Value = { member -> hex }
    /// </summary>
    public Dictionary<string, Dictionary<string, string>>? GroupedColorMappings { get; set; }
}

public class SizeEncoding
{
    public string Field { get; set; } = string.Empty;
    public string Aggregation { get; set; } = string.Empty;
}

public class TextEncoding
{
    public List<string> Fields { get; set; } = new();
}

public class ShapeEncoding
{
    public string Field { get; set; } = string.Empty;
    public string ShapeType { get; set; } = string.Empty; // filled/circle, square, etc.
}

/// <summary>
/// Represents a customized label (KPI display) structure with field roles
/// </summary>
public class CustomizedLabel
{
    /// <summary>
    /// Fields used in the label with their display roles
    /// </summary>
    public List<LabelFieldRole> FieldRoles { get; set; } = new();

    /// <summary>
    /// Raw formatted text content (for reference)
    /// </summary>
    public string RawContent { get; set; } = string.Empty;
}

/// <summary>
/// Represents a field's role in a customized label
/// </summary>
public class LabelFieldRole
{
    /// <summary>
    /// Display name of the field (or static text content if IsStaticText is true)
    /// </summary>
    public string FieldName { get; set; } = string.Empty;

    /// <summary>
    /// Role/purpose of the field in the label (e.g., "Primary Value", "Comparison Value", "Change Indicator")
    /// For static text, this will be "Static Text"
    /// </summary>
    public string Role { get; set; } = string.Empty;

    /// <summary>
    /// Formatting applied (bold, fontsize, color)
    /// </summary>
    public string Formatting { get; set; } = string.Empty;

    /// <summary>
    /// Indicates whether this element is static text (not a field reference)
    /// </summary>
    public bool IsStaticText { get; set; } = false;

    /// <summary>
    /// Font color hex code (e.g., #820000) for display in KPI table
    /// </summary>
    public string? FontColor { get; set; }

    /// <summary>
    /// Indicates whether the font color was inherited from the worksheet's default style
    /// (rather than being explicitly set on this element)
    /// </summary>
    public bool IsInherited { get; set; } = false;
}

/// <summary>
/// Contains map-specific configuration for geographic visualizations
/// </summary>
public class MapConfiguration
{
    /// <summary>
    /// Geographic fields used (Latitude/Longitude generated or custom)
    /// </summary>
    public List<GeographicField> GeographicFields { get; set; } = new();

    /// <summary>
    /// Map source provider (e.g., "Tableau", "Mapbox", "WMS")
    /// </summary>
    public string? MapSource { get; set; }

    /// <summary>
    /// Map layer configuration settings
    /// </summary>
    public MapLayerSettings? LayerSettings { get; set; }

    /// <summary>
    /// Base map style (e.g., "normal", "light", "dark")
    /// </summary>
    public string? BaseMapStyle { get; set; }

    /// <summary>
    /// Map washout level (0.0 to 1.0, where 1.0 is fully washed out/faded)
    /// </summary>
    public double? Washout { get; set; }

    /// <summary>
    /// Whether the map has a geometry encoding (spatial data)
    /// </summary>
    public bool HasGeometryEncoding { get; set; }

    /// <summary>
    /// Geometry column reference if present
    /// </summary>
    public string? GeometryColumn { get; set; }

    /// <summary>
    /// Zoom level or scale configuration
    /// </summary>
    public string? ZoomLevel { get; set; }

    /// <summary>
    /// Map projection type (e.g., "mercator", "equirectangular")
    /// </summary>
    public string? Projection { get; set; }

    /// <summary>
    /// Whether map labels are shown
    /// </summary>
    public bool? ShowLabels { get; set; }
}

/// <summary>
/// Represents a geographic field used in the map
/// </summary>
public class GeographicField
{
    /// <summary>
    /// Field name (display name)
    /// </summary>
    public string FieldName { get; set; } = string.Empty;

    /// <summary>
    /// Geographic role (e.g., "Latitude", "Longitude", "State", "Country", "ZipCode")
    /// </summary>
    public string GeographicRole { get; set; } = string.Empty;

    /// <summary>
    /// Whether this is a generated field (e.g., "[Latitude (generated)]")
    /// </summary>
    public bool IsGenerated { get; set; }

    /// <summary>
    /// The shelf where this field is used (Rows, Columns, Detail)
    /// </summary>
    public string Shelf { get; set; } = string.Empty;
}

/// <summary>
/// Map layer visibility and styling settings
/// </summary>
public class MapLayerSettings
{
    /// <summary>
    /// Whether country/territory borders are shown
    /// </summary>
    public bool? ShowCountryBorders { get; set; }

    /// <summary>
    /// Whether state/province borders are shown
    /// </summary>
    public bool? ShowStateBorders { get; set; }

    /// <summary>
    /// Whether county borders are shown
    /// </summary>
    public bool? ShowCountyBorders { get; set; }

    /// <summary>
    /// Whether city names are shown
    /// </summary>
    public bool? ShowCityNames { get; set; }

    /// <summary>
    /// Whether coastlines are shown
    /// </summary>
    public bool? ShowCoastlines { get; set; }

    /// <summary>
    /// Whether roads/streets are shown
    /// </summary>
    public bool? ShowStreets { get; set; }

    /// <summary>
    /// Whether water features (lakes, rivers) are shown
    /// </summary>
    public bool? ShowWaterFeatures { get; set; }

    /// <summary>
    /// Additional layer configurations as key-value pairs
    /// </summary>
    public Dictionary<string, string> AdditionalLayers { get; set; } = new();
}

/// <summary>
/// Contains table-specific configuration for table/crosstab visualizations
/// </summary>
public class TableConfiguration
{
    /// <summary>
    /// Table type (e.g., "Crosstab", "Text Table", "Highlight Table")
    /// </summary>
    public string TableType { get; set; } = string.Empty;

    /// <summary>
    /// Columns in the table with their display order
    /// </summary>
    public List<TableColumn> Columns { get; set; } = new();

    /// <summary>
    /// Dimensions displayed in rows
    /// </summary>
    public List<string> RowDimensions { get; set; } = new();

    /// <summary>
    /// Whether the table uses Measure Names construct (measures as columns)
    /// </summary>
    public bool UsesMeasureNames { get; set; }

    /// <summary>
    /// Visual formatting settings for the table
    /// </summary>
    public TableFormatting? Formatting { get; set; }

    /// <summary>
    /// Whether row banding is enabled
    /// </summary>
    public bool HasRowBanding { get; set; }

    /// <summary>
    /// Whether column banding is enabled
    /// </summary>
    public bool HasColumnBanding { get; set; }
}

/// <summary>
/// Represents a column in a table visualization
/// </summary>
public class TableColumn
{
    /// <summary>
    /// Display name shown in the column header
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Internal field reference
    /// </summary>
    public string InternalName { get; set; } = string.Empty;

    /// <summary>
    /// Aggregation function (SUM, AVG, COUNT, etc.)
    /// </summary>
    public string Aggregation { get; set; } = string.Empty;

    /// <summary>
    /// Number format pattern (e.g., "p0.0%", "c\"$\"#,##0")
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// Display order (0-based index)
    /// </summary>
    public int DisplayOrder { get; set; }

    /// <summary>
    /// Calculation name for custom aggregations (used for linking to calculation details)
    /// </summary>
    public string? CalculationName { get; set; }
}

/// <summary>
/// Visual formatting settings for table visualizations
/// </summary>
public class TableFormatting
{
    /// <summary>
    /// Header background color (hex code)
    /// </summary>
    public string? HeaderBackgroundColor { get; set; }

    /// <summary>
    /// Header text color (hex code)
    /// </summary>
    public string? HeaderTextColor { get; set; }

    /// <summary>
    /// Row banding color (hex code)
    /// </summary>
    public string? RowBandingColor { get; set; }

    /// <summary>
    /// Column banding color (hex code)
    /// </summary>
    public string? ColumnBandingColor { get; set; }

    /// <summary>
    /// Font family for the table
    /// </summary>
    public string? FontFamily { get; set; }

    /// <summary>
    /// Font size in points
    /// </summary>
    public int? FontSize { get; set; }

    /// <summary>
    /// Cell background color (hex code)
    /// </summary>
    public string? CellBackgroundColor { get; set; }

    /// <summary>
    /// Text alignment for row headers
    /// </summary>
    public string? RowHeaderAlignment { get; set; }

    /// <summary>
    /// Vertical alignment for row headers
    /// </summary>
    public string? RowHeaderVerticalAlignment { get; set; }
}
