using System.Xml.Linq;
using NeonScribe.Tableau.Core.Models;
using NeonScribe.Tableau.Core.Resolution;
using NeonScribe.Tableau.Core.Utilities;

namespace NeonScribe.Tableau.Core.Parsers;

public class DataSourceParser
{
    private readonly NameResolver _nameResolver;

    public DataSourceParser(NameResolver nameResolver)
    {
        _nameResolver = nameResolver;
    }

    public List<DataSource> ParseDataSources(XDocument document)
    {
        var dataSources = new List<DataSource>();

        var dsElements = document.Descendants("datasource")
            .Where(ds => ds.Attribute("inline")?.Value == "true" &&
                        !IsParameterDataSource(ds))
            .ToList();

        foreach (var dsElement in dsElements)
        {
            var dataSource = new DataSource
            {
                InternalName = dsElement.Attribute("name")?.Value ?? string.Empty,
                Caption = dsElement.Attribute("caption")?.Value ?? string.Empty,
                Name = dsElement.Attribute("caption")?.Value ?? string.Empty
            };

            // Parse connection type and details
            var connection = dsElement.Descendants("connection").FirstOrDefault();
            dataSource.ConnectionType = connection?.Attribute("class")?.Value ?? string.Empty;
            dataSource.ConnectionDetails = ExtractConnectionDetails(connection);

            // Parse custom SQL if present
            dataSource.CustomSql = ExtractCustomSql(dsElement);

            // Parse fields
            dataSource.Fields = ParseFields(dsElement);

            // Parse aliases
            dataSource.Aliases = ParseAliases(dsElement);

            dataSources.Add(dataSource);
        }

        return dataSources;
    }

    public List<Parameter> ParseParameters(XDocument document)
    {
        var parameters = new List<Parameter>();

        var paramDatasources = document.Descendants("datasource")
            .Where(IsParameterDataSource)
            .ToList();

        foreach (var paramDs in paramDatasources)
        {
            var columns = paramDs.Descendants("column").ToList();

            foreach (var column in columns)
            {
                var parameter = new Parameter
                {
                    InternalName = column.Attribute("name")?.Value ?? string.Empty,
                    Caption = _nameResolver.GetDisplayName(column),
                    Name = _nameResolver.GetDisplayName(column),
                    DataType = column.Attribute("datatype")?.Value ?? string.Empty,
                    DomainType = column.Attribute("param-domain-type")?.Value ?? string.Empty
                };

                // Parse current and default values
                var valueAttr = column.Attribute("value")?.Value;
                parameter.CurrentValue = ParseValue(valueAttr, parameter.DataType);

                var calculation = column.Element("calculation");
                var defaultFormula = calculation?.Attribute("formula")?.Value;
                parameter.DefaultValue = ParseValue(defaultFormula, parameter.DataType);

                // Parse value aliases (maps value -> display name)
                parameter.ValueAliases = ParseParameterValueAliases(column);

                // Set the default display name based on the default value and aliases
                // First check for 'alias' attribute on the column itself (direct default display name)
                var columnAlias = column.Attribute("alias")?.Value;
                if (!string.IsNullOrEmpty(columnAlias))
                {
                    parameter.DefaultDisplayName = columnAlias;
                }
                else if (parameter.DefaultValue != null && parameter.ValueAliases != null)
                {
                    // Try to look up by the raw value format (for datetime: #yyyy-MM-dd HH:mm:ss#)
                    var rawDefaultValue = defaultFormula?.Trim('"', '\'');
                    if (!string.IsNullOrEmpty(rawDefaultValue) && parameter.ValueAliases.TryGetValue(rawDefaultValue, out var displayName))
                    {
                        parameter.DefaultDisplayName = displayName;
                    }
                    else
                    {
                        // Fallback to ToString() for non-datetime types
                        var defaultKey = parameter.DefaultValue.ToString() ?? "";
                        if (parameter.ValueAliases.TryGetValue(defaultKey, out displayName))
                        {
                            parameter.DefaultDisplayName = displayName;
                        }
                    }
                }

                // Parse allowable values
                parameter.AllowableValues = ParseParameterAllowableValues(column, parameter.DataType);

                parameters.Add(parameter);
            }
        }

        return parameters;
    }

    private List<Field> ParseFields(XElement datasource)
    {
        var fields = new List<Field>();

        var columns = datasource.Descendants("column").ToList();

        foreach (var column in columns)
        {
            var field = new Field
            {
                InternalName = column.Attribute("name")?.Value ?? string.Empty,
                Name = _nameResolver.GetDisplayName(column),
                DataType = column.Attribute("datatype")?.Value ?? string.Empty,
                Role = column.Attribute("role")?.Value ?? string.Empty,
                Type = column.Attribute("type")?.Value ?? string.Empty
            };

            // Check if it's a calculated field
            var calculation = column.Element("calculation");
            if (calculation != null)
            {
                field.IsCalculated = true;
                var rawFormula = calculation.Attribute("formula")?.Value ?? string.Empty;
                // Replace internal field names with display names in formulas
                field.Formula = _nameResolver.ReplaceFieldReferencesInFormula(rawFormula);

                // Detect calculation type
                DetectCalculationType(field);

                // Extract dependencies from formula (using the display names)
                field.Dependencies = _nameResolver.ExtractFieldReferencesFromFormula(field.Formula);
            }

            fields.Add(field);
        }

        return fields;
    }

    private void DetectCalculationType(Field field)
    {
        if (string.IsNullOrEmpty(field.Formula))
            return;

        var formula = field.Formula;

        // Detect LOD calculations
        if (formula.Contains("{ FIXED"))
        {
            field.CalculationType = "LOD";
            field.LodType = "FIXED";
            field.LodDimensions = CalculationExplainer.ExtractLodDimensions(formula, "FIXED");
            field.Explanation = CalculationExplainer.ExplainLodCalculation(field);
        }
        else if (formula.Contains("{ INCLUDE"))
        {
            field.CalculationType = "LOD";
            field.LodType = "INCLUDE";
            field.LodDimensions = CalculationExplainer.ExtractLodDimensions(formula, "INCLUDE");
            field.Explanation = CalculationExplainer.ExplainLodCalculation(field);
        }
        else if (formula.Contains("{ EXCLUDE"))
        {
            field.CalculationType = "LOD";
            field.LodType = "EXCLUDE";
            field.LodDimensions = CalculationExplainer.ExtractLodDimensions(formula, "EXCLUDE");
            field.Explanation = CalculationExplainer.ExplainLodCalculation(field);
        }
        // Detect table calculations
        else if (formula.Contains("RUNNING_") || formula.Contains("WINDOW_") ||
                 formula.Contains("LOOKUP") || formula.Contains("INDEX()") ||
                 formula.Contains("RANK") || formula.Contains("PERCENT_OF_TOTAL"))
        {
            field.CalculationType = "Table Calculation";
            field.TableCalcFunction = CalculationExplainer.DetectTableCalcFunction(formula);
            field.Explanation = CalculationExplainer.ExplainTableCalculation(field);
        }
    }

    private List<AliasMapping> ParseAliases(XElement datasource)
    {
        var aliasMappings = new List<AliasMapping>();

        var columns = datasource.Descendants("column").ToList();

        foreach (var column in columns)
        {
            var aliasElements = column.Descendants("alias").ToList();

            if (aliasElements.Any())
            {
                var mapping = new AliasMapping
                {
                    Field = _nameResolver.GetDisplayName(column),
                    Mappings = new List<AliasEntry>()
                };

                foreach (var alias in aliasElements)
                {
                    var entry = new AliasEntry
                    {
                        Key = alias.Attribute("key")?.Value ?? string.Empty,
                        Value = alias.Attribute("value")?.Value ?? string.Empty
                    };

                    mapping.Mappings.Add(entry);
                }

                aliasMappings.Add(mapping);
            }
        }

        return aliasMappings;
    }

    /// <summary>
    /// Parse value aliases for a parameter (maps raw value to display name)
    /// e.g., 1 -> "Age Group", 2 -> "Prior Edu"
    /// </summary>
    private Dictionary<string, string>? ParseParameterValueAliases(XElement column)
    {
        var aliases = new Dictionary<string, string>();

        // Aliases can come from <aliases><alias key="1" value="Age Group" /></aliases>
        var aliasElements = column.Element("aliases")?.Elements("alias").ToList();
        if (aliasElements != null)
        {
            foreach (var alias in aliasElements)
            {
                var key = alias.Attribute("key")?.Value;
                var value = alias.Attribute("value")?.Value;
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                {
                    aliases[key] = value;
                }
            }
        }

        // Also check <members><member value="1" alias="Age Group" /></members>
        var memberElements = column.Element("members")?.Elements("member").ToList();
        if (memberElements != null)
        {
            foreach (var member in memberElements)
            {
                var value = member.Attribute("value")?.Value;
                var alias = member.Attribute("alias")?.Value;
                if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(alias) && !aliases.ContainsKey(value))
                {
                    aliases[value] = alias;
                }
            }
        }

        return aliases.Count > 0 ? aliases : null;
    }

    private ParameterAllowableValues? ParseParameterAllowableValues(XElement column, string dataType)
    {
        var domainType = column.Attribute("param-domain-type")?.Value;

        if (string.IsNullOrEmpty(domainType) || domainType == "all")
            return null;

        var allowableValues = new ParameterAllowableValues
        {
            Type = domainType
        };

        if (domainType == "range")
        {
            var range = column.Element("range");
            if (range != null)
            {
                allowableValues.Min = ParseValue(range.Attribute("min")?.Value, dataType);
                allowableValues.Max = ParseValue(range.Attribute("max")?.Value, dataType);
                allowableValues.Granularity = ParseValue(range.Attribute("granularity")?.Value, dataType);
            }
        }
        else if (domainType == "list")
        {
            var members = column.Descendants("member").ToList();
            var values = new List<object>();
            var displayNames = new List<string>();

            foreach (var member in members)
            {
                var value = ParseValue(member.Attribute("value")?.Value, dataType);
                var alias = member.Attribute("alias")?.Value;
                if (value != null)
                {
                    values.Add(value);
                    displayNames.Add(alias ?? value.ToString() ?? "");
                }
            }

            allowableValues.Values = values;
            allowableValues.DisplayNames = displayNames;
        }

        return allowableValues;
    }

    private object? ParseValue(string? value, string dataType)
    {
        if (string.IsNullOrEmpty(value))
            return null;

        // Remove quotes for strings
        value = value.Trim('"', '\'');

        return dataType switch
        {
            "integer" => int.TryParse(value, out var i) ? i : null,
            "real" => double.TryParse(value, out var d) ? d : null,
            "boolean" => bool.TryParse(value, out var b) ? b : null,
            "date" or "datetime" => DateTime.TryParse(value, out var dt) ? dt : null,
            _ => value
        };
    }

    private bool IsParameterDataSource(XElement datasource)
    {
        // Check for connection class="parameters" (older Tableau format)
        var connection = datasource.Descendants("connection").FirstOrDefault();
        if (connection?.Attribute("class")?.Value == "parameters")
            return true;

        // Check for name='Parameters' with hasconnection='false' (newer Tableau format)
        var name = datasource.Attribute("name")?.Value;
        var hasConnection = datasource.Attribute("hasconnection")?.Value;
        if (name == "Parameters" && hasConnection == "false")
            return true;

        return false;
    }

    private string? ExtractCustomSql(XElement datasource)
    {
        // Look for <relation type="text"> which contains custom SQL
        var textRelations = datasource.Descendants("relation")
            .Where(r => r.Attribute("type")?.Value == "text")
            .ToList();

        if (!textRelations.Any())
            return null;

        // The SQL is in the text content of the relation element
        var sql = textRelations.First().Value?.Trim();
        return string.IsNullOrWhiteSpace(sql) ? null : sql;
    }

    private string? ExtractConnectionDetails(XElement? connection)
    {
        if (connection == null)
            return null;

        var details = new List<string>();

        // Extract common connection attributes from the connection itself
        ExtractConnectionAttributes(connection, details);

        // Also check named connections (nested connections)
        var namedConnections = connection.Descendants("connection").ToList();
        foreach (var namedConn in namedConnections)
        {
            ExtractConnectionAttributes(namedConn, details);
        }

        return details.Any() ? string.Join(", ", details) : null;
    }

    private void ExtractConnectionAttributes(XElement connection, List<string> details)
    {
        var server = connection.Attribute("server")?.Value;
        if (!string.IsNullOrEmpty(server) && !details.Any(d => d.StartsWith("Server:")))
            details.Add($"Server: {server}");

        var database = connection.Attribute("dbname")?.Value ?? connection.Attribute("database")?.Value;
        if (!string.IsNullOrEmpty(database) && !details.Any(d => d.StartsWith("Database:")))
            details.Add($"Database: {database}");

        var filename = connection.Attribute("filename")?.Value;
        if (!string.IsNullOrEmpty(filename) && !details.Any(d => d.StartsWith("File:")))
            details.Add($"File: {filename}");

        var schema = connection.Attribute("schema")?.Value;
        if (!string.IsNullOrEmpty(schema) && !details.Any(d => d.StartsWith("Schema:")))
            details.Add($"Schema: {schema}");
    }
}
