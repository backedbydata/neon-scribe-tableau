using System.Text.RegularExpressions;
using System.Xml.Linq;
using NeonScribe.Tableau.Core.Models;

namespace NeonScribe.Tableau.Core.Resolution;

/// <summary>
/// Internal class to store field metadata during parsing
/// </summary>
internal class FieldMetadata
{
    public string InternalName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string? DataSourceName { get; set; }
    public string? DataSourceId { get; set; }
    public string? DataType { get; set; }
    public string? Role { get; set; }
    public bool IsCalculated { get; set; }
    public string? Formula { get; set; }
    public string? PhysicalTable { get; set; }
    public string? PhysicalColumn { get; set; }
    public List<string>? Dependencies { get; set; }
}

public class NameResolver
{
    private readonly Dictionary<string, string> _internalToDisplayName = new();
    private readonly Dictionary<string, Dictionary<string, string>> _fieldAliases = new();
    private readonly Dictionary<string, FieldMetadata> _fieldMetadata = new();
    private readonly Dictionary<string, Dictionary<string, string>> _columnMappings = new(); // datasource -> (field -> physical column)

    public NameResolver(XDocument twbDocument)
    {
        BuildNameLookup(twbDocument);
    }

    private void BuildNameLookup(XDocument doc)
    {
        // Parse all datasources to build name mappings
        var datasources = doc.Descendants("datasource").ToList();

        foreach (var datasource in datasources)
        {
            var dsName = datasource.Attribute("name")?.Value;
            var dsCaption = datasource.Attribute("caption")?.Value;

            // Parse column mappings (field -> physical column) from connection/cols
            ParseColumnMappings(datasource, dsName);

            // Parse all columns in this datasource
            var columns = datasource.Descendants("column").ToList();

            foreach (var column in columns)
            {
                var internalName = column.Attribute("name")?.Value;
                var caption = column.Attribute("caption")?.Value;

                if (!string.IsNullOrEmpty(internalName))
                {
                    // Use caption if available, otherwise use cleaned internal name
                    var displayName = caption ?? CleanInternalName(internalName);
                    _internalToDisplayName[internalName] = displayName;

                    // Also store fully qualified names
                    if (!string.IsNullOrEmpty(dsName))
                    {
                        var qualifiedName = $"[{dsName}].{internalName}";
                        _internalToDisplayName[qualifiedName] = displayName;
                    }

                    // Build field metadata for lineage
                    var metadata = new FieldMetadata
                    {
                        InternalName = internalName,
                        DisplayName = displayName,
                        DataSourceName = dsCaption ?? dsName,
                        DataSourceId = dsName,
                        DataType = column.Attribute("datatype")?.Value,
                        Role = column.Attribute("role")?.Value
                    };

                    // Check if it's a calculated field
                    var calculation = column.Element("calculation");
                    if (calculation != null)
                    {
                        metadata.IsCalculated = true;
                        metadata.Formula = calculation.Attribute("formula")?.Value;
                    }

                    // Get physical column mapping if available
                    if (!string.IsNullOrEmpty(dsName) && _columnMappings.TryGetValue(dsName, out var colMappings))
                    {
                        if (colMappings.TryGetValue(internalName, out var physicalMapping))
                        {
                            // Parse [Table].[Column] format
                            var (table, col) = ParsePhysicalMapping(physicalMapping);
                            metadata.PhysicalTable = table;
                            metadata.PhysicalColumn = col;
                        }
                    }

                    _fieldMetadata[internalName] = metadata;

                    // Also store with qualified name
                    if (!string.IsNullOrEmpty(dsName))
                    {
                        var qualifiedName = $"[{dsName}].{internalName}";
                        _fieldMetadata[qualifiedName] = metadata;
                    }

                    // Parse aliases for this field
                    ParseAliases(column, internalName);
                }
            }

            // Parse parameters (they're in a special datasource)
            var connectionClass = datasource.Descendants("connection")
                .FirstOrDefault()?.Attribute("class")?.Value;

            if (connectionClass == "parameters")
            {
                foreach (var column in columns)
                {
                    var internalName = column.Attribute("name")?.Value;
                    var caption = column.Attribute("caption")?.Value;

                    if (!string.IsNullOrEmpty(internalName))
                    {
                        var displayName = caption ?? CleanInternalName(internalName);
                        _internalToDisplayName[internalName] = displayName;
                    }
                }
            }
        }

        // Parse datasource-dependencies inside worksheets
        // These contain local field definitions with captions that may differ from or supplement the main datasource
        var dsDependencies = doc.Descendants("datasource-dependencies").ToList();
        foreach (var dsDep in dsDependencies)
        {
            var dsName = dsDep.Attribute("datasource")?.Value;
            var depColumns = dsDep.Elements("column").ToList();

            foreach (var column in depColumns)
            {
                var internalName = column.Attribute("name")?.Value;
                var caption = column.Attribute("caption")?.Value;

                if (!string.IsNullOrEmpty(internalName) && !string.IsNullOrEmpty(caption))
                {
                    // Only add if we have a caption (don't overwrite with cleaned name)
                    // This ensures worksheet-local captions are captured
                    if (!_internalToDisplayName.ContainsKey(internalName))
                    {
                        _internalToDisplayName[internalName] = caption;
                    }

                    // Also store fully qualified names
                    if (!string.IsNullOrEmpty(dsName))
                    {
                        var qualifiedName = $"[{dsName}].{internalName}";
                        if (!_internalToDisplayName.ContainsKey(qualifiedName))
                        {
                            _internalToDisplayName[qualifiedName] = caption;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Parse column mappings from the connection/cols section
    /// This maps Tableau field names to physical table.column references
    /// </summary>
    private void ParseColumnMappings(XElement datasource, string? dsName)
    {
        if (string.IsNullOrEmpty(dsName))
            return;

        var connection = datasource.Element("connection");
        var colsElement = connection?.Element("cols");

        if (colsElement == null)
            return;

        var mappings = new Dictionary<string, string>();

        foreach (var map in colsElement.Elements("map"))
        {
            var key = map.Attribute("key")?.Value;
            var value = map.Attribute("value")?.Value;

            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
            {
                mappings[key] = value;
            }
        }

        if (mappings.Any())
        {
            _columnMappings[dsName] = mappings;
        }
    }

    /// <summary>
    /// Parse a physical mapping like [Table].[Column] into table and column parts
    /// </summary>
    private (string? table, string? column) ParsePhysicalMapping(string mapping)
    {
        var match = Regex.Match(mapping, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (match.Success)
        {
            return (match.Groups[1].Value, match.Groups[2].Value);
        }
        return (null, mapping);
    }

    private void ParseAliases(XElement column, string fieldInternalName)
    {
        var aliasElements = column.Descendants("alias").ToList();

        if (aliasElements.Any())
        {
            var aliases = new Dictionary<string, string>();

            foreach (var alias in aliasElements)
            {
                var key = alias.Attribute("key")?.Value;
                var value = alias.Attribute("value")?.Value;

                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                {
                    aliases[key] = value;
                }
            }

            if (aliases.Any())
            {
                _fieldAliases[fieldInternalName] = aliases;
            }
        }
    }

    private string CleanInternalName(string internalName)
    {
        var cleaned = internalName;

        // Handle qualified field references: [datasource].[field]
        // Extract just the field part
        var qualifiedMatch = Regex.Match(internalName, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (qualifiedMatch.Success)
        {
            // Take the field part (second capture group)
            cleaned = qualifiedMatch.Groups[2].Value;
        }
        else
        {
            // Simple field reference - just remove brackets
            cleaned = internalName.Trim('[', ']');

            // Remove datasource prefixes like "federated.abc123."
            var dotIndex = cleaned.LastIndexOf('.');
            if (dotIndex > 0 && cleaned.Substring(0, dotIndex).Contains("federated"))
            {
                cleaned = cleaned.Substring(dotIndex + 1);
            }
        }

        // Handle field references with derivation prefixes and type suffixes
        // Format can be: prefix:FieldName:suffix or prefix:prefix2:FieldName:suffix:suffix2
        // Examples: none:patient_id:nk, usr:Calculation_123:qk, pcto:usr:Patient Id (copy)_123:qk:2
        if (cleaned.Contains(':'))
        {
            var parts = cleaned.Split(':');
            if (parts.Length >= 2)
            {
                // Known prefixes that are derivation indicators (not field names)
                var knownPrefixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "none", "sum", "avg", "min", "max", "count", "countd", "attr", "median", "stdev",
                    "yr", "qr", "mn", "wk", "dy", "hr", "wd", "md", "tdy", "tqr", "tmn",
                    "usr", "pcto", // usr = user-defined, pcto = percent of total
                };

                // Find the first part that isn't a known prefix - that's likely the field name
                string fieldName = cleaned; // fallback
                for (int i = 0; i < parts.Length; i++)
                {
                    var part = parts[i];
                    // Skip known prefixes
                    if (knownPrefixes.Contains(part))
                        continue;
                    // Skip known suffixes (usually short type indicators at end)
                    if (i == parts.Length - 1 && part.Length <= 3)
                        continue;
                    if (parts.Length > 2 && i == parts.Length - 2 && parts[parts.Length - 1].Length <= 3 && part.Length <= 3)
                        continue;
                    // This looks like the field name
                    fieldName = part;
                    break;
                }
                cleaned = fieldName;
            }
        }

        return cleaned;
    }

    public string GetDisplayName(string internalName)
    {
        if (string.IsNullOrEmpty(internalName))
        {
            return "[Unknown Field]";
        }

        // Try direct lookup first
        if (_internalToDisplayName.TryGetValue(internalName, out var displayName))
        {
            return displayName;
        }

        // Try without brackets
        var cleanName = internalName.Trim('[', ']');
        if (_internalToDisplayName.TryGetValue($"[{cleanName}]", out displayName))
        {
            return displayName;
        }

        // Clean the name (removes aggregation prefixes/suffixes) and try lookup
        var cleaned = CleanInternalName(internalName);
        if (_internalToDisplayName.TryGetValue($"[{cleaned}]", out displayName))
        {
            return displayName;
        }

        // Fall back to the cleaned name (without brackets)
        return cleaned;
    }

    /// <summary>
    /// Get display name for a filter parameter, including derivation prefix (Year of, Quarter of, etc.)
    /// Filter params have format: [datasource].[derivation:FieldName:type]
    /// </summary>
    public string GetFilterDisplayName(string filterParam)
    {
        if (string.IsNullOrEmpty(filterParam))
        {
            return "[Unknown Field]";
        }

        // Extract the field reference part: [datasource].[derivation:FieldName:type]
        // We need the derivation prefix and the field name
        var match = Regex.Match(filterParam, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (!match.Success)
        {
            // Try simple format without datasource
            return GetDisplayName(filterParam);
        }

        var fieldPart = match.Groups[2].Value; // e.g., "yr:Application Submit Date:ok"

        // Parse derivation prefix if present
        var derivationPrefix = "";
        var fieldName = fieldPart;

        if (fieldPart.Contains(':'))
        {
            var parts = fieldPart.Split(':');
            if (parts.Length >= 2)
            {
                var derivation = parts[0];
                fieldName = parts[1];

                // Map derivation codes to readable prefixes
                derivationPrefix = derivation switch
                {
                    "yr" => "Year of ",
                    "qr" => "Quarter of ",
                    "mn" => "Month of ",
                    "dy" => "Day of ",
                    "wk" => "Week of ",
                    "md" => "Month/Day of ",
                    "mdy" => "Month/Day/Year of ",
                    "sum" => "Sum of ",
                    "avg" => "Average of ",
                    "cnt" => "Count of ",
                    "cntd" => "Count Distinct of ",
                    "min" => "Min of ",
                    "max" => "Max of ",
                    "attr" => "",
                    "none" => "",
                    "usr" => "",
                    _ => ""
                };
            }
        }

        // Get the display name for the base field
        var baseDisplayName = GetDisplayName($"[{fieldName}]");

        // Combine derivation prefix with display name
        return derivationPrefix + baseDisplayName;
    }

    public string GetDisplayName(XElement element)
    {
        // Priority: caption > name > fallback
        var caption = element.Attribute("caption")?.Value;
        if (!string.IsNullOrEmpty(caption))
        {
            return caption;
        }

        var name = element.Attribute("name")?.Value;
        if (!string.IsNullOrEmpty(name))
        {
            return GetDisplayName(name);
        }

        return "[Unknown Field]";
    }

    public string ApplyAlias(string fieldInternalName, string value)
    {
        if (_fieldAliases.TryGetValue(fieldInternalName, out var aliases))
        {
            if (aliases.TryGetValue(value, out var aliasedValue))
            {
                return aliasedValue;
            }
        }

        return value;
    }

    public Dictionary<string, string> GetAliasesForField(string fieldInternalName)
    {
        return _fieldAliases.TryGetValue(fieldInternalName, out var aliases)
            ? aliases
            : new Dictionary<string, string>();
    }

    public List<string> ExtractFieldReferencesFromFormula(string formula)
    {
        // Extract field references from calculation formulas
        // Tableau uses patterns like [Field Name] for field references
        var fieldPattern = new Regex(@"\[([^\]]+)\]");
        var matches = fieldPattern.Matches(formula);

        var fields = new List<string>();
        foreach (Match match in matches)
        {
            var fieldRef = match.Groups[1].Value;

            // Skip Tableau functions and keywords
            if (!IsTableauFunction(fieldRef))
            {
                fields.Add(fieldRef);
            }
        }

        return fields.Distinct().ToList();
    }

    private bool IsTableauFunction(string text)
    {
        // Common Tableau functions and keywords to exclude
        var functions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "FIXED", "INCLUDE", "EXCLUDE", "IF", "THEN", "ELSE", "END",
            "CASE", "WHEN", "AND", "OR", "NOT", "NULL", "TRUE", "FALSE"
        };

        return functions.Contains(text);
    }

    public Dictionary<string, string> GetAllMappings()
    {
        return new Dictionary<string, string>(_internalToDisplayName);
    }

    /// <summary>
    /// Build a FieldLineage object from a filter param (e.g., [datasource].[derivation:FieldName:type])
    /// This provides the full tracing from display name to data source element
    /// </summary>
    public FieldLineage? BuildFieldLineage(string filterParam)
    {
        if (string.IsNullOrEmpty(filterParam))
            return null;

        var lineage = new FieldLineage
        {
            InternalName = filterParam
        };

        // Extract datasource and field parts: [datasource].[derivation:FieldName:type]
        var match = Regex.Match(filterParam, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (!match.Success)
        {
            // Simple format without datasource prefix
            lineage.DisplayName = GetDisplayName(filterParam);
            lineage.BaseFieldName = CleanInternalName(filterParam);
            return lineage;
        }

        var datasourceId = match.Groups[1].Value;
        var fieldPart = match.Groups[2].Value; // e.g., "yr:Application Submit Date:ok"

        lineage.DataSourceId = datasourceId;

        // Parse derivation and field name
        string? derivationCode = null;
        string fieldName = fieldPart;

        if (fieldPart.Contains(':'))
        {
            var parts = fieldPart.Split(':');
            if (parts.Length >= 2)
            {
                derivationCode = parts[0];
                fieldName = parts[1];

                // Map derivation codes to readable names
                lineage.Derivation = derivationCode switch
                {
                    "yr" => "Year",
                    "qr" => "Quarter",
                    "mn" => "Month",
                    "dy" => "Day",
                    "wk" => "Week",
                    "md" => "Month/Day",
                    "mdy" => "Month/Day/Year",
                    "sum" => "Sum",
                    "avg" => "Average",
                    "cnt" => "Count",
                    "cntd" => "Count Distinct",
                    "min" => "Min",
                    "max" => "Max",
                    "attr" => null,
                    "none" => null,
                    "usr" => null,
                    _ => null
                };
            }
        }

        lineage.BaseFieldName = fieldName;

        // Build the display name (with derivation prefix if applicable)
        var derivationPrefix = lineage.Derivation != null ? $"{lineage.Derivation} of " : "";
        var baseDisplayName = GetDisplayName($"[{fieldName}]");
        lineage.DisplayName = derivationPrefix + baseDisplayName;

        // Try to get metadata for the base field
        var baseFieldKey = $"[{fieldName}]";
        if (_fieldMetadata.TryGetValue(baseFieldKey, out var metadata))
        {
            lineage.DataSourceName = metadata.DataSourceName;
            lineage.DataType = metadata.DataType;
            lineage.Role = metadata.Role;
            lineage.IsCalculated = metadata.IsCalculated;
            lineage.PhysicalTable = metadata.PhysicalTable;
            lineage.PhysicalColumn = metadata.PhysicalColumn;

            if (metadata.IsCalculated)
            {
                lineage.Formula = ReplaceFieldReferencesInFormula(metadata.Formula ?? "");
                lineage.Dependencies = ExtractFieldReferencesFromFormula(lineage.Formula ?? "");
            }
        }
        else
        {
            // Try with fully qualified name
            var qualifiedKey = $"[{datasourceId}].[{fieldName}]";
            if (_fieldMetadata.TryGetValue(qualifiedKey, out metadata))
            {
                lineage.DataSourceName = metadata.DataSourceName;
                lineage.DataType = metadata.DataType;
                lineage.Role = metadata.Role;
                lineage.IsCalculated = metadata.IsCalculated;
                lineage.PhysicalTable = metadata.PhysicalTable;
                lineage.PhysicalColumn = metadata.PhysicalColumn;

                if (metadata.IsCalculated)
                {
                    lineage.Formula = ReplaceFieldReferencesInFormula(metadata.Formula ?? "");
                    lineage.Dependencies = ExtractFieldReferencesFromFormula(lineage.Formula ?? "");
                }
            }
        }

        // If we still don't have datasource name, look it up from any field with that datasource ID
        if (string.IsNullOrEmpty(lineage.DataSourceName))
        {
            var anyFieldWithDs = _fieldMetadata.Values.FirstOrDefault(m => m.DataSourceId == datasourceId);
            if (anyFieldWithDs != null)
            {
                lineage.DataSourceName = anyFieldWithDs.DataSourceName;
            }
        }

        return lineage;
    }

    /// <summary>
    /// Get the derivation display name for a derivation code
    /// </summary>
    public string? GetDerivationDisplayName(string derivationCode)
    {
        return derivationCode switch
        {
            "yr" => "Year",
            "qr" => "Quarter",
            "mn" => "Month",
            "dy" => "Day",
            "wk" => "Week",
            "md" => "Month/Day",
            "mdy" => "Month/Day/Year",
            "sum" => "Sum",
            "avg" => "Average",
            "cnt" => "Count",
            "cntd" => "Count Distinct",
            "min" => "Min",
            "max" => "Max",
            _ => null
        };
    }

    public string ReplaceFieldReferencesInFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula))
        {
            return formula;
        }

        // Find all field references in the formula like [Field_Name_123]
        var fieldPattern = new Regex(@"\[([^\]]+)\]");
        var result = fieldPattern.Replace(formula, match =>
        {
            var fieldRef = match.Value; // e.g., "[Patient Id (copy)_48695209524043779]"

            // Skip Tableau functions and keywords
            var innerValue = match.Groups[1].Value;
            if (IsTableauFunction(innerValue))
            {
                return fieldRef; // Keep as-is
            }

            // Try to resolve to display name
            var displayName = GetDisplayName(fieldRef);

            // If we got a different name (not the cleaned version), use it
            // Otherwise keep the original reference
            return $"[{displayName}]";
        });

        return result;
    }
}
