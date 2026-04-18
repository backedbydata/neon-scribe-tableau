using System.Xml.Linq;
using NeonScribe.Tableau.Core.Models;
using NeonScribe.Tableau.Core.Resolution;

namespace NeonScribe.Tableau.Core.Parsers;

public class WorksheetParser
{
    private readonly NameResolver _nameResolver;

    public WorksheetParser(NameResolver nameResolver)
    {
        _nameResolver = nameResolver;
    }

    public List<Worksheet> ParseWorksheets(XDocument document)
    {
        var worksheets = new List<Worksheet>();

        // Build a set of hidden worksheet names from window elements
        var hiddenWorksheetNames = GetHiddenWorksheetNames(document);

        // Build global color mappings from datasource-level encodings
        var globalColorMappings = BuildGlobalColorMappings(document);

        // Build field member lookup from manual-sort definitions (used for grouping CASE color mappings)
        var fieldSortMembers = BuildFieldSortMembers(document);

        // Build CASE formula lookup: field internal name -> formula string
        var caseFormulas = BuildCaseFormulaLookup(document);

        // Build parameter alias lookup: param internal name -> { whenKey -> friendly label }
        var paramAliasLookup = BuildParameterAliasLookup(document);

        var wsElements = document.Descendants("worksheet").ToList();

        foreach (var wsElement in wsElements)
        {
            var worksheetName = wsElement.Attribute("name")?.Value ?? string.Empty;
            var worksheet = new Worksheet
            {
                Name = worksheetName,
                Caption = wsElement.Attribute("caption")?.Value ?? worksheetName,
                IsHidden = hiddenWorksheetNames.Contains(worksheetName)
            };

            // Parse mark type from <pane><mark class='...' /> element
            var panes = wsElement.Descendants("panes").FirstOrDefault();
            var markElement = panes?.Descendants("pane")
                .SelectMany(p => p.Elements("mark"))
                .FirstOrDefault(m => m.Attribute("class") != null);

            worksheet.MarkType = markElement?.Attribute("class")?.Value ?? "Automatic";
            worksheet.VisualType = DetermineVisualType(wsElement, worksheet.MarkType);

            // Parse fields used (from rows, cols, and encodings)
            worksheet.FieldsUsed = ParseFieldsUsed(wsElement);

            // Parse filters from worksheet element (worksheet-specific filters only)
            worksheet.Filters = ParseFilters(wsElement);

            // Parse tooltip
            worksheet.Tooltip = ParseTooltip(wsElement);

            // Parse mark encodings (Color, Size, Text, Shape, Detail, Tooltip)
            worksheet.MarkEncodings = ParseMarkEncodings(wsElement, globalColorMappings, fieldSortMembers, caseFormulas, paramAliasLookup);

            // Parse customized label for KPI displays
            worksheet.CustomizedLabel = ParseCustomizedLabel(wsElement);

            // Check if there's actual tooltip content (not just mark encoding fields)
            worksheet.HasActualTooltipContent = HasActualTooltipContent(wsElement);

            // Parse map configuration if this is a map visualization
            worksheet.MapConfiguration = ParseMapConfiguration(wsElement);

            // Parse title from layout-options if present
            worksheet.Title = ParseTitle(wsElement);

            // Parse table configuration if this is a table visualization
            worksheet.TableConfiguration = ParseTableConfiguration(document, wsElement, worksheet.VisualType);

            worksheets.Add(worksheet);
        }

        // NOTE: Shared-view filters are now parsed separately and associated with dashboards
        // See ParseSharedViewFilters() method which returns filters for dashboard association

        return worksheets;
    }

    /// <summary>
    /// Build a lookup of field names to color mappings from datasource-level style rules.
    /// Color encodings at the datasource level apply globally to worksheets using those fields.
    /// </summary>
    private Dictionary<string, Dictionary<string, string>> BuildGlobalColorMappings(XDocument document)
    {
        var mappings = new Dictionary<string, Dictionary<string, string>>();

        // Find all datasource elements and extract color encodings from their style rules
        var datasources = document.Descendants("datasource").ToList();

        foreach (var datasource in datasources)
        {
            var styleRules = datasource.Descendants("style-rule")
                .Where(sr => sr.Attribute("element")?.Value == "mark")
                .ToList();

            foreach (var styleRule in styleRules)
            {
                var colorEncodings = styleRule.Elements("encoding")
                    .Where(e => e.Attribute("attr")?.Value == "color")
                    .ToList();

                foreach (var encoding in colorEncodings)
                {
                    var field = encoding.Attribute("field")?.Value;
                    if (string.IsNullOrEmpty(field))
                        continue;

                    // Extract color mappings
                    var colorMap = new Dictionary<string, string>();
                    foreach (var map in encoding.Elements("map"))
                    {
                        var color = map.Attribute("to")?.Value;
                        var bucket = map.Element("bucket")?.Value;
                        if (!string.IsNullOrEmpty(color) && !string.IsNullOrEmpty(bucket))
                        {
                            var cleanBucket = bucket.Trim('"');
                            if (!colorMap.ContainsKey(cleanBucket))
                            {
                                colorMap[cleanBucket] = color;
                            }
                        }
                    }

                    if (colorMap.Any())
                    {
                        // Store with the field name (may need to normalize for lookup)
                        mappings[field] = colorMap;
                    }
                }
            }
        }

        return mappings;
    }

    /// <summary>
    /// Builds a lookup from field column name (e.g., "[none:Age Group:nk]") to its ordered member values
    /// extracted from manual-sort dictionary elements in the workbook.
    /// </summary>
    private Dictionary<string, List<string>> BuildFieldSortMembers(XDocument document)
    {
        var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var sort in document.Descendants("manual-sort"))
        {
            var column = sort.Attribute("column")?.Value;
            if (string.IsNullOrEmpty(column))
                continue;

            var members = sort.Descendants("bucket")
                .Select(b => b.Value.Trim('"'))
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();

            if (members.Any() && !result.ContainsKey(column))
                result[column] = members;
        }

        return result;
    }

    /// <summary>
    /// Builds a lookup from calculation internal name to CASE formula string.
    /// Only includes columns whose formula starts with "CASE".
    /// </summary>
    private Dictionary<string, string> BuildCaseFormulaLookup(XDocument document)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var column in document.Descendants("column"))
        {
            var name = column.Attribute("name")?.Value;
            if (string.IsNullOrEmpty(name))
                continue;

            var formula = column.Element("calculation")?.Attribute("formula")?.Value;
            if (!string.IsNullOrEmpty(formula) &&
                formula.TrimStart().StartsWith("CASE", StringComparison.OrdinalIgnoreCase))
            {
                if (!result.ContainsKey(name))
                    result[name] = formula;
            }
        }

        return result;
    }

    /// <summary>
    /// Builds a lookup from parameter internal name (e.g., "[Parameter 1]") to its value aliases
    /// (e.g., {"1": "Total", "2": "Race", "3": "Sex", "4": "Age"}).
    /// </summary>
    private Dictionary<string, Dictionary<string, string>> BuildParameterAliasLookup(XDocument document)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var column in document.Descendants("column"))
        {
            var domainType = column.Attribute("param-domain-type")?.Value;
            if (string.IsNullOrEmpty(domainType))
                continue;

            var name = column.Attribute("name")?.Value;
            if (string.IsNullOrEmpty(name))
                continue;

            var aliases = new Dictionary<string, string>();

            // <aliases><alias key="1" value="Total" /></aliases>
            foreach (var alias in column.Element("aliases")?.Elements("alias") ?? Enumerable.Empty<XElement>())
            {
                var key = alias.Attribute("key")?.Value;
                var value = alias.Attribute("value")?.Value;
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value) && !aliases.ContainsKey(key))
                    aliases[key] = value;
            }

            // <members><member value="1" alias="Total" /></members>
            foreach (var member in column.Element("members")?.Elements("member") ?? Enumerable.Empty<XElement>())
            {
                var value = member.Attribute("value")?.Value;
                var alias = member.Attribute("alias")?.Value;
                if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(alias) && !aliases.ContainsKey(value))
                    aliases[value] = alias;
            }

            if (aliases.Any() && !result.ContainsKey(name))
                result[name] = aliases;
        }

        return result;
    }

    /// <summary>
    /// If the color raw field references a CASE-based calculated field, groups color mappings
    /// by the CASE branches (parameter selection values). Uses manual-sort members to assign
    /// color mapping keys to the correct field branch group.
    /// Returns null if grouping is not applicable or produces only one group.
    /// </summary>
    private Dictionary<string, Dictionary<string, string>>? BuildGroupedColorMappings(
        Dictionary<string, string> colorMappings,
        string rawField,
        Dictionary<string, List<string>> fieldSortMembers,
        Dictionary<string, string> caseFormulas,
        Dictionary<string, Dictionary<string, string>> paramAliasLookup)
    {
        // Extract calc name from raw field: [none:Calculation_789818813289242634:nk] -> Calculation_789818813289242634
        var calcNameMatch = System.Text.RegularExpressions.Regex.Match(rawField, @":([^:]+):");
        var calcName = calcNameMatch.Success ? calcNameMatch.Groups[1].Value : rawField.Trim('[', ']');

        // Find CASE formula for this calc
        string? caseFormula = null;
        foreach (var (key, formula) in caseFormulas)
        {
            if (key.Contains(calcName, StringComparison.OrdinalIgnoreCase))
            {
                caseFormula = formula;
                break;
            }
        }

        if (string.IsNullOrEmpty(caseFormula))
            return null;

        // Find the controlling parameter and its aliases for friendly group labels
        // e.g., CASE [Parameters].[Parameter 1] -> look up "[Parameter 1]" in paramAliasLookup
        Dictionary<string, string>? paramAliases = null;
        var paramMatch = System.Text.RegularExpressions.Regex.Match(
            caseFormula, @"CASE\s+\[Parameters\]\.\[([^\]]+)\]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (paramMatch.Success)
        {
            var paramInternalName = $"[{paramMatch.Groups[1].Value}]";
            paramAliasLookup.TryGetValue(paramInternalName, out paramAliases);
        }

        // Parse WHEN/THEN branches
        var whenMatches = System.Text.RegularExpressions.Regex.Matches(
            caseFormula,
            @"WHEN\s+(\S+)\s+THEN\s+(.+?)(?=\s+WHEN|\s+END|$)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);

        if (whenMatches.Count == 0)
            return null;

        // Build group for each branch using friendly labels resolved immediately
        var result = new Dictionary<string, Dictionary<string, string>>();
        var claimedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var unresolvedBranchLabels = new List<string>(); // friendly labels for branches with no sort data

        foreach (System.Text.RegularExpressions.Match m in whenMatches)
        {
            var whenKey = m.Groups[1].Value.Trim();
            var then = m.Groups[2].Value.Trim();

            // Resolve friendly label from param aliases, fall back to whenKey
            var groupLabel = paramAliases != null && paramAliases.TryGetValue(whenKey, out var alias)
                ? alias
                : whenKey;

            if (then.StartsWith("\"") && then.EndsWith("\""))
            {
                // Literal: WHEN 1 THEN "Total"
                var literal = then.Trim('"');
                if (colorMappings.ContainsKey(literal))
                {
                    result[groupLabel] = new Dictionary<string, string> { [literal] = colorMappings[literal] };
                    claimedKeys.Add(literal);
                }
            }
            else
            {
                // Field reference: WHEN 2 THEN [Race Name]
                var fieldName = then.Trim('[', ']');

                // Find matching sort members
                List<string>? members = null;
                foreach (var (sortKey, sortValues) in fieldSortMembers)
                {
                    if (sortKey.Contains(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        members = sortValues;
                        break;
                    }
                }

                if (members != null)
                {
                    var groupEntries = new Dictionary<string, string>();
                    foreach (var member in members)
                    {
                        if (colorMappings.TryGetValue(member, out var hex) && !claimedKeys.Contains(member))
                        {
                            groupEntries[member] = hex;
                            claimedKeys.Add(member);
                        }
                    }
                    if (groupEntries.Any())
                        result[groupLabel] = groupEntries;
                    else
                        unresolvedBranchLabels.Add(groupLabel);
                }
                else
                {
                    // No sort data for this field — track for unclaimed assignment
                    unresolvedBranchLabels.Add(groupLabel);
                }
            }
        }

        // Assign unclaimed color mapping keys to unresolved branches using a combined label
        var unclaimed = colorMappings
            .Where(kv => !claimedKeys.Contains(kv.Key) && kv.Key != "%null%")
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        if (unclaimed.Any())
        {
            var combinedLabel = unresolvedBranchLabels.Any()
                ? string.Join(" / ", unresolvedBranchLabels)
                : "Other";
            result[combinedLabel] = unclaimed;
        }

        return result.Count > 1 ? result : null;
    }

    /// <summary>
    /// Look up color mappings for a field in the global mappings.
    /// Handles field name variations (with/without datasource prefix).
    /// </summary>
    private Dictionary<string, string>? LookupGlobalColorMappings(string fieldRef,
        Dictionary<string, Dictionary<string, string>> globalMappings)
    {
        // Direct match
        if (globalMappings.TryGetValue(fieldRef, out var directMatch))
        {
            return directMatch;
        }

        // Extract the field part after the datasource prefix (e.g., "[federated.xxx].[none:Field:nk]" -> "[none:Field:nk]")
        var fieldPart = ExtractFieldPart(fieldRef);
        if (!string.IsNullOrEmpty(fieldPart))
        {
            if (globalMappings.TryGetValue(fieldPart, out var partMatch))
            {
                return partMatch;
            }
        }

        // Try matching by field part in keys
        foreach (var (key, value) in globalMappings)
        {
            var keyFieldPart = ExtractFieldPart(key);
            if (keyFieldPart == fieldPart)
            {
                return value;
            }
        }

        return null;
    }

    /// <summary>
    /// Extract the field part from a fully-qualified field reference.
    /// E.g., "[federated.xxx].[none:Field:nk]" -> "[none:Field:nk]"
    /// </summary>
    private string? ExtractFieldPart(string fieldRef)
    {
        if (string.IsNullOrEmpty(fieldRef))
            return null;

        // Pattern: [datasource].[field] - extract the field part
        var lastDotBracket = fieldRef.LastIndexOf("].[");
        if (lastDotBracket > 0)
        {
            return fieldRef.Substring(lastDotBracket + 2); // Skip "]."
        }

        // If no datasource prefix, return as-is
        return fieldRef;
    }

    private List<FieldUsageInSheet> ParseFieldsUsed(XElement worksheet)
    {
        var fieldsUsed = new List<FieldUsageInSheet>();
        var table = worksheet.Element("table");

        if (table == null)
            return fieldsUsed;

        // Parse rows
        var rowsElement = table.Element("rows");
        if (rowsElement != null)
        {
            var fieldRefs = ExtractFieldReferences(rowsElement.Value);
            foreach (var fieldRef in fieldRefs)
            {
                fieldsUsed.Add(new FieldUsageInSheet
                {
                    Field = _nameResolver.GetDisplayName(fieldRef),
                    Shelf = "Rows",
                    Aggregation = ExtractAggregation(fieldRef)
                });
            }
        }

        // Parse columns
        var colsElement = table.Element("cols");
        if (colsElement != null)
        {
            var fieldRefs = ExtractFieldReferences(colsElement.Value);
            foreach (var fieldRef in fieldRefs)
            {
                fieldsUsed.Add(new FieldUsageInSheet
                {
                    Field = _nameResolver.GetDisplayName(fieldRef),
                    Shelf = "Columns",
                    Aggregation = ExtractAggregation(fieldRef)
                });
            }
        }

        // Parse color, size, shape, etc. from panes
        var panes = worksheet.Descendants("panes").FirstOrDefault();
        if (panes != null)
        {
            ParseEncodings(panes, fieldsUsed, "Color");
            ParseEncodings(panes, fieldsUsed, "Size");
            ParseEncodings(panes, fieldsUsed, "Shape");
        }

        return fieldsUsed;
    }

    private void ParseEncodings(XElement panes, List<FieldUsageInSheet> fieldsUsed, string encodingType)
    {
        var encodingElements = panes.Descendants("encoding")
            .Where(e => e.Attribute("attr")?.Value?.Equals(encodingType, StringComparison.OrdinalIgnoreCase) == true);

        foreach (var encoding in encodingElements)
        {
            var fieldAttr = encoding.Attribute("field")?.Value;
            if (!string.IsNullOrEmpty(fieldAttr))
            {
                fieldsUsed.Add(new FieldUsageInSheet
                {
                    Field = _nameResolver.GetDisplayName(fieldAttr),
                    Shelf = encodingType,
                    Aggregation = ExtractAggregation(fieldAttr)
                });
            }
        }
    }

    private List<string> ExtractFieldReferences(string value)
    {
        // Field references are in the format [datasource].[field]
        var fields = new List<string>();

        // Use regex to properly extract full field references like [datasource].[field]
        // This pattern matches: [anything].[anything:anything:anything]
        var fullRefRegex = new System.Text.RegularExpressions.Regex(@"\[([^\]]+)\]\.\[([^\]]+)\]");
        var fullMatches = fullRefRegex.Matches(value);

        foreach (System.Text.RegularExpressions.Match match in fullMatches)
        {
            // Only add the field part (group 2), not the datasource part (group 1)
            fields.Add($"[{match.Groups[2].Value}]");
        }

        // If no full references found, fall back to simple bracket extraction
        // This handles cases like [Field Name] without datasource prefix
        if (fields.Count == 0)
        {
            var simpleRegex = new System.Text.RegularExpressions.Regex(@"\[([^\]]+)\]");
            var matches = simpleRegex.Matches(value);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                fields.Add(match.Value);
            }
        }

        return fields;
    }

    private string ExtractAggregation(string fieldRef)
    {
        // Extract aggregation from field references like [sum:Sales:qk]
        if (fieldRef.Contains(":"))
        {
            var parts = fieldRef.Trim('[', ']').Split(':');
            if (parts.Length >= 2)
            {
                var agg = parts[0];
                return agg.ToUpper() switch
                {
                    "SUM" => "SUM",
                    "AVG" => "AVG",
                    "COUNT" => "COUNT",
                    "MIN" => "MIN",
                    "MAX" => "MAX",
                    "COUNTD" => "COUNTD",
                    "NONE" => "None",
                    _ => "None"
                };
            }
        }

        return "None";
    }

    private List<Filter> ParseFilters(XElement worksheet)
    {
        var filters = new List<Filter>();
        var filterElements = worksheet.Descendants("filter").ToList();

        foreach (var filterElement in filterElements)
        {
            var columnValue = filterElement.Attribute("column")?.Value ?? string.Empty;

            // Skip internal action filters - these are Tableau's internal filter actions,
            // not user-visible filters. They have field names like "[Action (Source,Target)]"
            // The column format may be "[datasource].[Action (field)]" or just "[Action (field)]"
            if (columnValue.Contains("[Action (") || columnValue.Contains("[Action("))
                continue;

            var filter = new Filter
            {
                Field = _nameResolver.GetDisplayName(columnValue),
                Type = filterElement.Attribute("class")?.Value ?? string.Empty
            };

            // Also skip if the resolved display name indicates an action filter
            // (in case the name resolver returns a cleaned name with Action prefix)
            if (filter.Field.StartsWith("Action (") || filter.Field.StartsWith("Action(") ||
                filter.Field.StartsWith("[Action (") || filter.Field.StartsWith("[Action("))
                continue;

            // Determine filter type and control type
            filter.FilterType = MapFilterClass(filter.Type);
            filter.ControlType = DetermineControlType(filterElement);

            // Parse type-specific attributes
            if (filter.Type == "quantitative")
            {
                ParseQuantitativeFilter(filterElement, filter);
            }
            else if (filter.Type == "categorical")
            {
                ParseCategoricalFilter(filterElement, filter);
            }
            else if (filter.Type.Contains("date"))
            {
                ParseDateFilter(filterElement, filter);
            }

            filters.Add(filter);
        }

        return filters;
    }

    private void ParseQuantitativeFilter(XElement filterElement, Filter filter)
    {
        var minElement = filterElement.Element("min");
        var maxElement = filterElement.Element("max");

        filter.Min = minElement != null ? double.Parse(minElement.Value) : null;
        filter.Max = maxElement != null ? double.Parse(maxElement.Value) : null;
        filter.IncludeNullValues = filterElement.Attribute("included-values")?.Value != "non-null";
    }

    private void ParseCategoricalFilter(XElement filterElement, Filter filter)
    {
        // Check for multi-select
        var groupFilter = filterElement.Element("groupfilter");
        if (groupFilter != null)
        {
            var ns = groupFilter.GetNamespaceOfPrefix("user");
            if (ns != null)
            {
                filter.AllowMultipleValues = groupFilter.Attribute(ns + "ui-enumeration")?.Value == "exclusive";
                filter.ShowOnlyRelevantValues = groupFilter.Attribute(ns + "ui-domain")?.Value == "relevant";
            }

            // Extract selected/default values from groupfilter
            var selectedValues = ExtractGroupFilterValues(groupFilter);
            if (selectedValues.Any())
            {
                filter.DefaultSelection = string.Join(", ", selectedValues);
            }
        }
    }

    /// <summary>
    /// Extract selected values from a groupfilter element.
    /// Handles single member selection, union of members, and except (exclusion) patterns.
    /// </summary>
    private List<string> ExtractGroupFilterValues(XElement groupFilter)
    {
        var values = new List<string>();
        var function = groupFilter.Attribute("function")?.Value;

        if (function == "member")
        {
            // Single member selection: <groupfilter function='member' member='2025' .../>
            var member = groupFilter.Attribute("member")?.Value;
            if (!string.IsNullOrEmpty(member))
            {
                values.Add(CleanFilterValue(member));
            }
        }
        else if (function == "union")
        {
            // Union of multiple members: <groupfilter function='union'><groupfilter function='member' .../></groupfilter>
            foreach (var child in groupFilter.Elements("groupfilter"))
            {
                values.AddRange(ExtractGroupFilterValues(child));
            }
        }
        else if (function == "except")
        {
            // Exclusion pattern: <groupfilter function='except'>
            //   <groupfilter function='level-members' level='...'/>  (all values)
            //   <groupfilter function='union'>...</groupfilter>      (values to exclude)
            // </groupfilter>
            var excludedValues = new List<string>();
            var children = groupFilter.Elements("groupfilter").ToList();
            foreach (var child in children)
            {
                var childFunction = child.Attribute("function")?.Value;
                if (childFunction == "union" || childFunction == "member")
                {
                    excludedValues.AddRange(ExtractGroupFilterValues(child));
                }
            }
            if (excludedValues.Any())
            {
                values.Add($"(All except: {string.Join(", ", excludedValues)})");
            }
        }
        else if (function == "level-members")
        {
            // All values selected (or combined with except for exclusion)
            // The level attribute shows which dimension level
            var level = groupFilter.Attribute("level")?.Value;
            if (!string.IsNullOrEmpty(level))
            {
                // Only add "(All)" if this is a standalone filter, not part of 'except'
                var parent = groupFilter.Parent;
                if (parent?.Name.LocalName != "groupfilter" ||
                    parent?.Attribute("function")?.Value != "except")
                {
                    values.Add("(All)");
                }
            }
        }

        return values;
    }

    /// <summary>
    /// Clean a filter value by removing surrounding quotes and normalizing the display
    /// </summary>
    private string CleanFilterValue(string value)
    {
        // Remove surrounding quotes if present
        if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length >= 2)
        {
            value = value.Substring(1, value.Length - 2);
        }
        return value;
    }

    private void ParseDateFilter(XElement filterElement, Filter filter)
    {
        filter.Period = filterElement.Element("period")?.Value;
        filter.Anchor = filterElement.Element("anchor")?.Value;

        var offsetElement = filterElement.Element("period-offset");
        if (offsetElement != null && int.TryParse(offsetElement.Value, out var offset))
        {
            filter.PeriodOffset = offset;
        }
    }

    private string MapFilterClass(string filterClass)
    {
        return filterClass switch
        {
            "categorical" => "Categorical (List)",
            "quantitative" => "Range",
            "relative-date-filter" => "Relative Date",
            _ => filterClass
        };
    }

    private string DetermineControlType(XElement filterElement)
    {
        var filterClass = filterElement.Attribute("class")?.Value;

        return filterClass switch
        {
            "categorical" => "Multi-select dropdown",
            "quantitative" => "Range slider",
            "relative-date-filter" => "Date range",
            _ => "Unknown"
        };
    }

    private Tooltip? ParseTooltip(XElement worksheet)
    {
        // First try the modern <customized-tooltip> element
        var customizedTooltip = worksheet.Descendants("customized-tooltip").FirstOrDefault();
        if (customizedTooltip != null)
        {
            var formattedText = customizedTooltip.Element("formatted-text");
            if (formattedText != null)
            {
                // Extract text content from <run> elements
                var runs = formattedText.Elements("run");
                var tooltipParts = new List<string>();
                var fieldsUsed = new List<string>();

                foreach (var run in runs)
                {
                    var text = run.Value;

                    // Handle line breaks - Tableau uses 'Æ' character as a line break marker
                    // It can appear as "Æ", "Æ ", "Æ\n", or with surrounding whitespace
                    var trimmedText = text.Trim();
                    if (trimmedText == "Æ" || trimmedText == "Æ\n" || text == "\n" ||
                        text.Trim('\n', '\r', ' ', '\t') == "Æ")
                    {
                        tooltipParts.Add("\n");
                        continue;
                    }

                    // Remove any remaining Æ characters that are used as formatting markers
                    text = text.Replace("Æ", "").Replace("\n", " ").Trim();

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // Replace field references with friendly names
                    var resolvedText = ResolveFieldReferencesInText(text, fieldsUsed);
                    tooltipParts.Add(resolvedText);
                }

                // Join parts and clean up whitespace around newlines
                var content = string.Join("", tooltipParts)
                    .Replace(" \n ", "\n")
                    .Replace(" \n", "\n")
                    .Replace("\n ", "\n")
                    .Trim();

                if (!string.IsNullOrEmpty(content))
                {
                    return new Tooltip
                    {
                        HasCustomTooltip = true,
                        Content = content,
                        FieldsUsed = fieldsUsed.Distinct().ToList()
                    };
                }
            }
        }

        // Fall back to legacy <format attr="tooltip"> approach
        var formatElements = worksheet.Descendants("format")
            .Where(f => f.Attribute("attr")?.Value == "tooltip")
            .ToList();

        if (!formatElements.Any())
            return null;

        var tooltipFormat = formatElements.First();
        var legacyContent = tooltipFormat.Attribute("value")?.Value ?? string.Empty;
        var legacyFieldsUsed = new List<string>();
        var resolvedLegacyContent = ResolveFieldReferencesInText(legacyContent, legacyFieldsUsed);

        return new Tooltip
        {
            HasCustomTooltip = !string.IsNullOrEmpty(resolvedLegacyContent),
            Content = resolvedLegacyContent,
            FieldsUsed = legacyFieldsUsed.Distinct().ToList()
        };
    }

    /// <summary>
    /// Replace field references like &lt;[datasource].[field:qk]&gt; with friendly display names wrapped in brackets
    /// </summary>
    private string ResolveFieldReferencesInText(string text, List<string> fieldsUsed)
    {
        // Pattern to match <[datasource].[field]> references
        var fieldRefPattern = new System.Text.RegularExpressions.Regex(@"<(\[[^\]]+\]\.\[[^\]]+\])>");

        return fieldRefPattern.Replace(text, match =>
        {
            var fieldRef = match.Groups[1].Value;
            var displayName = _nameResolver.GetDisplayName(fieldRef);
            fieldsUsed.Add(displayName);
            // Wrap field references in brackets to distinguish from plain text
            return $"[{displayName}]";
        });
    }

    private MarkEncodings? ParseMarkEncodings(XElement worksheet,
        Dictionary<string, Dictionary<string, string>> globalColorMappings,
        Dictionary<string, List<string>> fieldSortMembers,
        Dictionary<string, string> caseFormulas,
        Dictionary<string, Dictionary<string, string>> paramAliasLookup)
    {
        var panes = worksheet.Descendants("panes").FirstOrDefault();
        if (panes == null)
            return null;

        var encodings = new MarkEncodings();
        var hasAnyEncoding = false;
        string? colorFieldForLookup = null;

        // Parse color encoding
        var colorEncoding = panes.Descendants("encoding")
            .FirstOrDefault(e => e.Attribute("attr")?.Value == "color");
        if (colorEncoding != null)
        {
            var field = colorEncoding.Attribute("field")?.Value;
            if (!string.IsNullOrEmpty(field))
            {
                colorFieldForLookup = field;
                encodings.Color = new ColorEncoding
                {
                    Field = _nameResolver.GetDisplayName(field),
                    RawField = field,
                    PaletteType = colorEncoding.Attribute("type")?.Value ?? "palette",
                    PaletteName = colorEncoding.Attribute("palette")?.Value
                };

                // Parse color mappings from <map to='#color'><bucket>value</bucket></map>
                foreach (var map in colorEncoding.Elements("map"))
                {
                    var color = map.Attribute("to")?.Value;
                    var bucket = map.Element("bucket")?.Value;
                    if (!string.IsNullOrEmpty(color) && !string.IsNullOrEmpty(bucket))
                    {
                        // Clean up bucket value (remove quotes)
                        var cleanBucket = bucket.Trim('"');
                        if (!encodings.Color.ColorMappings.ContainsKey(cleanBucket))
                        {
                            encodings.Color.ColorMappings[cleanBucket] = color;
                        }
                    }
                }

                hasAnyEncoding = true;
            }
        }

        // Also check for color from column attribute (alternate encoding format)
        var colorColumn = panes.Descendants("encodings").FirstOrDefault()?
            .Elements("color").FirstOrDefault();
        if (colorColumn != null && encodings.Color == null)
        {
            var field = colorColumn.Attribute("column")?.Value;
            if (!string.IsNullOrEmpty(field))
            {
                colorFieldForLookup = field;
                encodings.Color = new ColorEncoding
                {
                    Field = _nameResolver.GetDisplayName(field),
                    RawField = field,
                    PaletteType = "column"
                };
                hasAnyEncoding = true;
            }
        }

        // If we have a color encoding but no color mappings, look up in global mappings
        if (encodings.Color != null && !encodings.Color.ColorMappings.Any() && colorFieldForLookup != null)
        {
            var mappings = LookupGlobalColorMappings(colorFieldForLookup, globalColorMappings);
            if (mappings != null)
            {
                foreach (var (key, value) in mappings)
                {
                    encodings.Color.ColorMappings[key] = value;
                }
            }
        }

        // If we have a color encoding with mappings and a raw field, try to group by CASE formula branches
        if (encodings.Color != null && encodings.Color.ColorMappings.Any() && !string.IsNullOrEmpty(encodings.Color.RawField))
        {
            encodings.Color.GroupedColorMappings = BuildGroupedColorMappings(
                encodings.Color.ColorMappings, encodings.Color.RawField, fieldSortMembers, caseFormulas, paramAliasLookup);
        }

        // Parse size encoding
        var sizeColumn = panes.Descendants("encodings").FirstOrDefault()?
            .Elements("size").FirstOrDefault();
        if (sizeColumn != null)
        {
            var field = sizeColumn.Attribute("column")?.Value;
            if (!string.IsNullOrEmpty(field))
            {
                encodings.Size = new SizeEncoding
                {
                    Field = _nameResolver.GetDisplayName(field),
                    Aggregation = ExtractAggregation(field)
                };
                hasAnyEncoding = true;
            }
        }

        // Parse text/label encoding
        var textColumns = panes.Descendants("encodings").FirstOrDefault()?
            .Elements("text").ToList();
        if (textColumns != null && textColumns.Any())
        {
            encodings.Text = new TextEncoding();
            foreach (var textCol in textColumns)
            {
                var field = textCol.Attribute("column")?.Value;
                if (!string.IsNullOrEmpty(field))
                {
                    encodings.Text.Fields.Add(_nameResolver.GetDisplayName(field));
                }
            }
            if (encodings.Text.Fields.Any())
            {
                hasAnyEncoding = true;
            }
        }

        // Parse tooltip fields from encoding
        var tooltipColumns = panes.Descendants("encodings").FirstOrDefault()?
            .Elements("tooltip").ToList();
        if (tooltipColumns != null && tooltipColumns.Any())
        {
            foreach (var tooltipCol in tooltipColumns)
            {
                var field = tooltipCol.Attribute("column")?.Value;
                if (!string.IsNullOrEmpty(field))
                {
                    encodings.TooltipFields.Add(_nameResolver.GetDisplayName(field));
                }
            }
            if (encodings.TooltipFields.Any())
            {
                hasAnyEncoding = true;
            }
        }

        // Parse shape from style-rule
        var shapeFormat = panes.Descendants("style-rule")
            .Where(sr => sr.Attribute("element")?.Value == "mark")
            .SelectMany(sr => sr.Elements("format"))
            .FirstOrDefault(f => f.Attribute("attr")?.Value == "shape");
        if (shapeFormat != null)
        {
            var shapeValue = shapeFormat.Attribute("value")?.Value;
            if (!string.IsNullOrEmpty(shapeValue))
            {
                encodings.Shape = new ShapeEncoding
                {
                    ShapeType = shapeValue
                };
                hasAnyEncoding = true;
            }
        }

        // Parse mark-color from style-rule (the primary visualization color)
        // Note: style-rule elements are in worksheet/style, not inside panes
        var markStyleRules = worksheet.Descendants("style-rule")
            .Where(sr => sr.Attribute("element")?.Value == "mark");

        var markColorFormat = markStyleRules
            .SelectMany(sr => sr.Elements("format"))
            .FirstOrDefault(f => f.Attribute("attr")?.Value == "mark-color");
        if (markColorFormat != null)
        {
            var markColorValue = markColorFormat.Attribute("value")?.Value;
            if (!string.IsNullOrEmpty(markColorValue))
            {
                // Ensure we have a Color encoding to attach mark-color to
                if (encodings.Color == null)
                {
                    encodings.Color = new ColorEncoding();
                }
                encodings.Color.MarkColor = markColorValue;
                hasAnyEncoding = true;
            }
        }

        // If we have a color encoding with a field but no mark-color, try to extract from color-palette
        // This handles gradient/sequential palettes where the last color represents the "high" value
        if (encodings.Color != null && !string.IsNullOrEmpty(encodings.Color.Field) && string.IsNullOrEmpty(encodings.Color.MarkColor))
        {
            var colorPalette = markStyleRules
                .Descendants("encoding")
                .Where(e => e.Attribute("attr")?.Value == "color")
                .Descendants("color-palette")
                .FirstOrDefault();
            if (colorPalette != null)
            {
                // Get the last color from the palette (typically the "high" end of the gradient)
                var lastColor = colorPalette.Elements("color").LastOrDefault()?.Value;
                if (!string.IsNullOrEmpty(lastColor))
                {
                    encodings.Color.MarkColor = lastColor;
                }
            }
        }

        return hasAnyEncoding ? encodings : null;
    }

    /// <summary>
    /// Determines the visual type based on mark type and worksheet structure.
    /// Detects KPI/Big Number displays when mark is Automatic but has text encodings and customized-label.
    /// Detects Map visualizations based on geographic fields, map sources, and map-specific style rules.
    /// </summary>
    private string DetermineVisualType(XElement worksheet, string markType)
    {
        var markTypeLower = markType.ToLowerInvariant();

        // If mark type is explicit (not Automatic), use the standard mapping
        if (markTypeLower != "automatic")
        {
            return MapMarkTypeToVisualType(markType);
        }

        // For "Automatic" mark type, analyze the worksheet structure to determine actual visual type
        var panes = worksheet.Descendants("panes").FirstOrDefault();
        if (panes == null)
            return "Automatic";

        var encodings = panes.Descendants("encodings").FirstOrDefault();
        var table = worksheet.Element("table");
        var view = table?.Parent;

        // Check for Map visualization indicators (check this FIRST before other types)
        if (IsMapVisualization(worksheet, table, panes))
        {
            return "Map";
        }

        // Check for KPI/Big Number indicator:
        // - Has customized-label (used for KPI text display)
        // - Has at least one text encoding
        // - Has no rows/cols OR rows/cols contain only dimensions (no measures)
        var hasCustomizedLabel = panes.Descendants("customized-label").Any();
        var textEncodingCount = encodings?.Elements("text").Count() ?? 0;
        var hasRows = !string.IsNullOrWhiteSpace(table?.Element("rows")?.Value);
        var hasCols = !string.IsNullOrWhiteSpace(table?.Element("cols")?.Value);

        if (hasCustomizedLabel && textEncodingCount >= 1)
        {
            // Check if rows/cols contain only dimensions (not measures)
            var rowsOnlyDimensions = !hasRows || !ContainsMeasureField(table?.Element("rows")?.Value ?? "");
            var colsOnlyDimensions = !hasCols || !ContainsMeasureField(table?.Element("cols")?.Value ?? "");

            if (rowsOnlyDimensions && colsOnlyDimensions)
            {
                return "KPI / Big Number";
            }
        }

        // Check for simple text table: has text encoding but no customized label, no rows/cols
        if (textEncodingCount > 0 && !hasRows && !hasCols)
        {
            return "Text / Label";
        }

        // Check if it's really just showing text labels (mark-labels-show = true without data axes)
        var markLabelsShow = panes.Descendants("format")
            .Any(f => f.Attribute("attr")?.Value == "mark-labels-show" && f.Attribute("value")?.Value == "true");

        if (markLabelsShow && textEncodingCount > 0 && !hasRows && !hasCols)
        {
            return "Text / Label";
        }

        // For Automatic mark type with rows/cols, infer visual type from axis structure
        // Tableau's Automatic typically resolves to bar chart when there's a dimension on one axis
        // and a measure on the other
        if (hasRows && hasCols)
        {
            var inferredType = InferVisualTypeFromAxes(table!);
            if (inferredType != null)
                return inferredType;
        }

        // Check for text displays with dimension rows only (no columns)
        // This includes decorative elements, labels, and spacers
        if (hasRows && !hasCols)
        {
            var rowsValue = table?.Element("rows")?.Value ?? "";
            if (!ContainsMeasureField(rowsValue))
            {
                return "Text / Label";
            }
        }

        return "Automatic";
    }

    /// <summary>
    /// Infer the visual type from the rows/columns structure.
    /// - Dimension on rows + [:Measure Names] on columns = Crosstab Table
    /// - Dimension on rows + measure on columns = Horizontal Bar Chart or Text Table
    /// - Measure on rows + dimension on columns = Vertical Bar Chart
    /// - Multiple dimensions without measures = Text Table
    /// </summary>
    private string? InferVisualTypeFromAxes(XElement table)
    {
        var rowsValue = table.Element("rows")?.Value ?? string.Empty;
        var colsValue = table.Element("cols")?.Value ?? string.Empty;

        // Check for [:Measure Names] construct - indicates a crosstab/pivot table
        if (colsValue.Contains("[:Measure Names]"))
        {
            return "Crosstab Table";
        }

        if (rowsValue.Contains("[:Measure Names]"))
        {
            return "Crosstab Table";
        }

        // Extract field types from the references
        var rowHasMeasure = ContainsMeasureField(rowsValue);
        var colHasMeasure = ContainsMeasureField(colsValue);
        var rowHasDimension = ContainsDimensionField(rowsValue);
        var colHasDimension = ContainsDimensionField(colsValue);

        // Nested dimension/measure in rows with measure in columns (e.g., "[dimension] / [measure]")
        // This is a horizontal bar chart with breakdowns
        if (rowHasDimension && rowHasMeasure && colHasMeasure && !colHasDimension)
        {
            return "Horizontal Bar";
        }

        // Dimension on rows with measures on columns - could be table or bar chart
        // If mark type is text/automatic, check for visual encodings to determine the type
        if (rowHasDimension && colHasMeasure && !rowHasMeasure)
        {
            // Check if it's a table by looking at mark type and encodings
            var panesElement = table.Parent?.Descendants("panes").FirstOrDefault();
            var markElement = panesElement?
                .Descendants("pane")
                .SelectMany(p => p.Elements("mark"))
                .FirstOrDefault(m => m.Attribute("class") != null);

            var markType = markElement?.Attribute("class")?.Value?.ToLowerInvariant() ?? "automatic";

            // Check for color encodings - if present, it's likely a bar chart
            var hasColorEncoding = panesElement?.Descendants("encodings")
                .Any(enc => enc.Element("color") != null) ?? false;

            if (markType == "automatic" || markType == "text")
            {
                // If there are color encodings, treat as horizontal bar chart
                if (hasColorEncoding)
                {
                    return "Horizontal Bar";
                }
                return "Text Table";
            }

            return "Horizontal Bar";
        }

        // Measure on rows, dimension on columns = Vertical Bar
        if (colHasDimension && rowHasMeasure && !colHasMeasure)
        {
            return "Vertical Bar";
        }

        // Both have dimensions, no measures = Text Table
        if (rowHasDimension && colHasDimension && !rowHasMeasure && !colHasMeasure)
        {
            return "Text Table";
        }

        return null;
    }

    /// <summary>
    /// Check if the field reference string contains measure fields.
    /// Measures are identified by aggregation prefixes (sum:, avg:, min:, max:, cnt:) or :qk suffix.
    /// </summary>
    private bool ContainsMeasureField(string fieldRef)
    {
        if (string.IsNullOrEmpty(fieldRef))
            return false;

        // Common measure patterns in Tableau field references
        // Note: usr: can be either measure or dimension depending on the suffix
        return fieldRef.Contains("sum:") ||
               fieldRef.Contains("avg:") ||
               fieldRef.Contains("min:") ||
               fieldRef.Contains("max:") ||
               fieldRef.Contains("cnt:") ||
               fieldRef.Contains("countd:") ||
               fieldRef.Contains(":qk");     // :qk indicates quantitative key (may have suffix like :qk:3])
    }

    /// <summary>
    /// Check if the field reference string contains dimension fields.
    /// Dimensions are primarily identified by the none: prefix, which is the most reliable
    /// indicator of unaggregated dimension fields.
    /// Note: :nk and :ok suffixes can appear on both dimensions AND measures (e.g., sum:Field:ok),
    /// so we focus on none: which is exclusive to dimensions.
    /// </summary>
    private bool ContainsDimensionField(string fieldRef)
    {
        if (string.IsNullOrEmpty(fieldRef))
            return false;

        // none: prefix always indicates a dimension
        return fieldRef.Contains("none:");
    }

    private string MapMarkTypeToVisualType(string markType)
    {
        return markType.ToLowerInvariant() switch
        {
            "bar" => "Bar Chart",
            "line" => "Line Chart",
            "area" => "Area Chart",
            "pie" => "Pie Chart",
            "circle" => "Circle/Bubble Chart",
            "square" => "Square Chart",
            "shape" => "Shape Chart",
            "text" => "Text Table",
            "map" => "Map",
            "polygon" => "Polygon/Map",
            "ganttbar" => "Gantt Chart",
            "automatic" => "Automatic",
            _ => markType
        };
    }

    /// <summary>
    /// Parse customized-label element to extract field roles for KPI displays.
    /// Analyzes the formatted-text structure to determine each field's role.
    /// </summary>
    private CustomizedLabel? ParseCustomizedLabel(XElement worksheet)
    {
        var panes = worksheet.Descendants("panes").FirstOrDefault();
        var customizedLabel = panes?.Descendants("customized-label").FirstOrDefault();
        if (customizedLabel == null)
            return null;

        var formattedText = customizedLabel.Element("formatted-text");
        if (formattedText == null)
            return null;

        // Extract default text color from worksheet style (for inheriting when no explicit fontcolor)
        // This is in <style><style-rule element='cell'><format attr='color' value='#xxx'>
        var defaultColor = worksheet.Descendants("style-rule")
            .Where(sr => sr.Attribute("element")?.Value == "cell")
            .SelectMany(sr => sr.Elements("format"))
            .FirstOrDefault(f => f.Attribute("attr")?.Value == "color")
            ?.Attribute("value")?.Value;

        var label = new CustomizedLabel();
        var runs = formattedText.Elements("run").ToList();
        var fieldRoles = new List<LabelFieldRole>();
        var fieldIndex = 0;

        // Track static text context for role determination
        var lastStaticText = "";

        foreach (var run in runs)
        {
            var content = run.Value;
            var isBold = run.Attribute("bold")?.Value == "true";
            var fontsize = run.Attribute("fontsize")?.Value;
            var fontcolor = run.Attribute("fontcolor")?.Value;

            // Determine effective color and whether it's inherited
            var effectiveColor = fontcolor;
            var isInherited = false;
            if (string.IsNullOrEmpty(fontcolor) && !string.IsNullOrEmpty(defaultColor))
            {
                effectiveColor = defaultColor;
                isInherited = true;
            }

            // Check if this run contains field references (may have multiple in one run)
            if (content.Contains("<[") && content.Contains("]>"))
            {
                // Extract ALL field references from CDATA or direct content
                var fieldRefMatches = System.Text.RegularExpressions.Regex.Matches(content, @"<\[([^\]]+)\]\.\[([^\]]+)\]>");
                foreach (System.Text.RegularExpressions.Match fieldRefMatch in fieldRefMatches)
                {
                    var fieldRef = $"[{fieldRefMatch.Groups[2].Value}]";
                    var displayName = _nameResolver.GetDisplayName(fieldRef);

                    // Determine role based on position, formatting, and context
                    var role = DetermineFieldRole(fieldIndex, displayName, isBold, fontsize, fontcolor, lastStaticText);

                    var formatting = new List<string>();
                    if (isBold) formatting.Add("bold");
                    if (!string.IsNullOrEmpty(fontsize)) formatting.Add($"size:{fontsize}");
                    if (!string.IsNullOrEmpty(effectiveColor)) formatting.Add($"color:{effectiveColor}");

                    fieldRoles.Add(new LabelFieldRole
                    {
                        FieldName = displayName,
                        Role = role,
                        Formatting = string.Join(", ", formatting),
                        IsStaticText = false,
                        FontColor = effectiveColor,
                        IsInherited = isInherited
                    });

                    fieldIndex++;
                }
            }
            else if (!string.IsNullOrWhiteSpace(content) && !content.Contains("Æ"))
            {
                // Static text - add as a label element and track for context
                var staticText = content.Trim();
                lastStaticText = staticText;

                // Add static text as a label element (for display in documentation)
                var formatting = new List<string>();
                if (isBold) formatting.Add("bold");
                if (!string.IsNullOrEmpty(fontsize)) formatting.Add($"size:{fontsize}");
                if (!string.IsNullOrEmpty(effectiveColor)) formatting.Add($"color:{effectiveColor}");

                fieldRoles.Add(new LabelFieldRole
                {
                    FieldName = $"\"{staticText}\"",
                    Role = "Static Text",
                    Formatting = string.Join(", ", formatting),
                    IsStaticText = true,
                    FontColor = effectiveColor,
                    IsInherited = isInherited
                });
            }
        }

        if (fieldRoles.Any())
        {
            label.FieldRoles = fieldRoles;
            label.RawContent = formattedText.ToString();
            return label;
        }

        return null;
    }

    /// <summary>
    /// Determine the role of a field in a KPI label based on context
    /// </summary>
    private string DetermineFieldRole(int index, string fieldName, bool isBold, string? fontsize, string? fontcolor, string lastStaticText)
    {
        var fieldLower = fieldName.ToLowerInvariant();

        // Check field name patterns
        if (fieldLower.Contains("current") || fieldLower.Contains("- current"))
            return "Primary Value (Current Period)";

        if (fieldLower.Contains("- pm") || fieldLower.Contains("prior month") || fieldLower.Contains("previous"))
            return "Comparison Value (Prior Period)";

        if (fieldLower.Contains("% change") || fieldLower.Contains("percent"))
        {
            if (fieldLower.Contains("+ve") || fieldLower.Contains("positive"))
                return "Percentage Change (Positive)";
            if (fieldLower.Contains("-ve") || fieldLower.Contains("negative"))
                return "Percentage Change (Negative)";
            return "Percentage Change";
        }

        if (fieldLower.Contains("arrow"))
        {
            if (fieldLower.Contains("+ve") || fieldLower.Contains("positive"))
                return "Trend Indicator (▲ Up)";
            if (fieldLower.Contains("-ve") || fieldLower.Contains("negative"))
                return "Trend Indicator (▼ Down)";
            return "Trend Indicator";
        }

        // Check context from preceding static text
        if (lastStaticText.ToLowerInvariant().Contains("vs pm"))
            return "Comparison Indicator";

        // Check formatting patterns
        if (index == 0 && isBold && !string.IsNullOrEmpty(fontsize))
        {
            if (int.TryParse(fontsize, out var size) && size >= 18)
                return "Primary Value (Headline)";
        }

        // Default based on position
        return index == 0 ? "Primary Value" : $"Supporting Value #{index}";
    }

    /// <summary>
    /// Check if the worksheet has actual tooltip content (not just mark encoding fields)
    /// </summary>
    private bool HasActualTooltipContent(XElement worksheet)
    {
        var panes = worksheet.Descendants("panes").FirstOrDefault();
        var customizedTooltip = panes?.Descendants("customized-tooltip").FirstOrDefault();

        if (customizedTooltip == null)
            return false;

        var formattedText = customizedTooltip.Element("formatted-text");
        if (formattedText == null)
            return false;

        // Check if there's any actual content (runs with text or field references)
        var runs = formattedText.Elements("run").ToList();
        foreach (var run in runs)
        {
            var content = run.Value;
            if (!string.IsNullOrWhiteSpace(content))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Determines if a worksheet is a map visualization based on multiple indicators.
    /// Checks for: map sources, geographic fields (Latitude/Longitude), geometry encodings, and map-specific style rules.
    /// </summary>
    private bool IsMapVisualization(XElement worksheet, XElement? table, XElement? panes)
    {
        if (table == null || panes == null)
            return false;

        var view = table.Parent;
        if (view == null)
            return false;

        // Indicator 1: Check for <mapsources> element in the view
        var hasMapSources = view.Elements("mapsources").Any();
        if (hasMapSources)
            return true;

        // Indicator 2: Check for generated geographic fields (Latitude/Longitude)
        var rowsValue = table.Element("rows")?.Value ?? string.Empty;
        var colsValue = table.Element("cols")?.Value ?? string.Empty;
        var hasLatitude = rowsValue.Contains("Latitude (generated)") || rowsValue.Contains("Latitude(generated)");
        var hasLongitude = colsValue.Contains("Longitude (generated)") || colsValue.Contains("Longitude(generated)");
        if (hasLatitude && hasLongitude)
            return true;

        // Indicator 3: Check for geometry encoding in panes
        var hasGeometryEncoding = panes.Elements("geometry").Any();
        if (hasGeometryEncoding)
            return true;

        // Indicator 4: Check for space encoding with geographic fields (alternative lat/long encoding)
        var spaceEncodings = panes.Descendants("encoding")
            .Where(e => e.Attribute("attr")?.Value == "space")
            .ToList();
        if (spaceEncodings.Count >= 2)
        {
            // Check if the fields are latitude/longitude
            foreach (var encoding in spaceEncodings)
            {
                var field = encoding.Attribute("field")?.Value ?? string.Empty;
                if (field.Contains("Latitude") || field.Contains("Longitude"))
                    return true;
            }
        }

        // Indicator 5: Check for map-specific style rules
        var hasMapLayerStyle = worksheet.Descendants("style-rule")
            .Any(sr => sr.Attribute("element")?.Value == "map-layer");
        if (hasMapLayerStyle)
            return true;

        var hasMapStyle = worksheet.Descendants("style-rule")
            .Any(sr => sr.Attribute("element")?.Value == "map");
        if (hasMapStyle)
            return true;

        return false;
    }

    /// <summary>
    /// Parse map-specific configuration from a worksheet.
    /// Extracts geographic fields, map sources, layer settings, and styling information.
    /// </summary>
    private MapConfiguration? ParseMapConfiguration(XElement worksheet)
    {
        var table = worksheet.Element("table");
        if (table == null)
            return null;

        var view = table.Parent;
        var panes = worksheet.Descendants("panes").FirstOrDefault();

        if (view == null || panes == null)
            return null;

        // Only parse if this is actually a map
        if (!IsMapVisualization(worksheet, table, panes))
            return null;

        var mapConfig = new MapConfiguration();

        // Parse geographic fields from rows/cols
        mapConfig.GeographicFields = ParseGeographicFields(table, panes);

        // Parse map source
        var mapSource = view.Elements("mapsources")
            .Elements("mapsource")
            .FirstOrDefault()?
            .Attribute("name")?.Value;
        mapConfig.MapSource = mapSource;

        // Parse geometry encoding
        var geometryElement = panes.Elements("geometry").FirstOrDefault();
        if (geometryElement != null)
        {
            mapConfig.HasGeometryEncoding = true;
            mapConfig.GeometryColumn = geometryElement.Attribute("column")?.Value;
            if (!string.IsNullOrEmpty(mapConfig.GeometryColumn))
            {
                mapConfig.GeometryColumn = _nameResolver.GetDisplayName(mapConfig.GeometryColumn);
            }
        }

        // Parse map layer settings
        mapConfig.LayerSettings = ParseMapLayerSettings(worksheet);

        // Parse map washout (fade level)
        var mapStyleRule = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "map");
        if (mapStyleRule != null)
        {
            var washoutFormat = mapStyleRule.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "washout");
            if (washoutFormat != null && double.TryParse(washoutFormat.Attribute("value")?.Value, out var washout))
            {
                mapConfig.Washout = washout;
            }
        }

        // Parse map style
        var baseStyleFormat = worksheet.Descendants("format")
            .FirstOrDefault(f => f.Attribute("attr")?.Value == "map-style");
        if (baseStyleFormat != null)
        {
            mapConfig.BaseMapStyle = baseStyleFormat.Attribute("value")?.Value;
        }

        // Check if mark labels are shown
        var markLabelsShow = panes.Descendants("format")
            .FirstOrDefault(f => f.Attribute("attr")?.Value == "mark-labels-show");
        if (markLabelsShow != null && bool.TryParse(markLabelsShow.Attribute("value")?.Value, out var showLabels))
        {
            mapConfig.ShowLabels = showLabels;
        }

        return mapConfig;
    }

    /// <summary>
    /// Parse geographic fields from rows, columns, and encodings
    /// </summary>
    private List<GeographicField> ParseGeographicFields(XElement table, XElement panes)
    {
        var geoFields = new List<GeographicField>();

        // Parse from rows
        var rowsValue = table.Element("rows")?.Value ?? string.Empty;
        AddGeographicFieldsFromAxisValue(rowsValue, "Rows", geoFields);

        // Parse from columns
        var colsValue = table.Element("cols")?.Value ?? string.Empty;
        AddGeographicFieldsFromAxisValue(colsValue, "Columns", geoFields);

        // Parse from space encodings (latitude/longitude)
        var spaceEncodings = panes.Descendants("encoding")
            .Where(e => e.Attribute("attr")?.Value == "space")
            .ToList();
        foreach (var encoding in spaceEncodings)
        {
            var fieldRef = encoding.Attribute("field")?.Value;
            if (!string.IsNullOrEmpty(fieldRef))
            {
                var displayName = _nameResolver.GetDisplayName(fieldRef);
                var role = DetermineGeographicRole(displayName, fieldRef);
                var isGenerated = fieldRef.Contains("(generated)") || displayName.Contains("(generated)");

                geoFields.Add(new GeographicField
                {
                    FieldName = displayName,
                    GeographicRole = role,
                    IsGenerated = isGenerated,
                    Shelf = "Space"
                });
            }
        }

        return geoFields;
    }

    /// <summary>
    /// Add geographic fields from an axis value (rows or cols)
    /// </summary>
    private void AddGeographicFieldsFromAxisValue(string axisValue, string shelf, List<GeographicField> geoFields)
    {
        if (string.IsNullOrEmpty(axisValue))
            return;

        var fieldRefs = ExtractFieldReferences(axisValue);
        foreach (var fieldRef in fieldRefs)
        {
            var displayName = _nameResolver.GetDisplayName(fieldRef);
            var role = DetermineGeographicRole(displayName, fieldRef);

            // Only add if this is actually a geographic field
            if (role != "Unknown")
            {
                var isGenerated = fieldRef.Contains("(generated)") || displayName.Contains("(generated)");

                geoFields.Add(new GeographicField
                {
                    FieldName = displayName,
                    GeographicRole = role,
                    IsGenerated = isGenerated,
                    Shelf = shelf
                });
            }
        }
    }

    /// <summary>
    /// Determine the geographic role based on field name
    /// </summary>
    private string DetermineGeographicRole(string displayName, string fieldRef)
    {
        var lowerName = displayName.ToLowerInvariant();
        var lowerRef = fieldRef.ToLowerInvariant();

        if (lowerName.Contains("latitude") || lowerRef.Contains("latitude"))
            return "Latitude";
        if (lowerName.Contains("longitude") || lowerRef.Contains("longitude"))
            return "Longitude";
        if (lowerName.Contains("country") || lowerRef.Contains("country"))
            return "Country";
        if (lowerName.Contains("state") || lowerRef.Contains("state"))
            return "State/Province";
        if (lowerName.Contains("city") || lowerRef.Contains("city"))
            return "City";
        if (lowerName.Contains("zip") || lowerName.Contains("postal") || lowerRef.Contains("zip") || lowerRef.Contains("postal"))
            return "Postal Code";
        if (lowerName.Contains("county") || lowerRef.Contains("county"))
            return "County";
        if (lowerName.Contains("region") || lowerRef.Contains("region"))
            return "Region";

        return "Unknown";
    }

    /// <summary>
    /// Parse map layer settings from style rules
    /// </summary>
    private MapLayerSettings? ParseMapLayerSettings(XElement worksheet)
    {
        var mapLayerStyleRule = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "map-layer");

        if (mapLayerStyleRule == null)
            return null;

        var settings = new MapLayerSettings();
        var formats = mapLayerStyleRule.Elements("format").ToList();

        // Map common layer format IDs to settings
        foreach (var format in formats)
        {
            var id = format.Attribute("id")?.Value;
            var value = format.Attribute("value")?.Value;
            var enabled = format.Attribute("enabled")?.Value;

            if (string.IsNullOrEmpty(id))
                continue;

            // Parse boolean value
            bool? boolValue = null;
            if (!string.IsNullOrEmpty(enabled))
            {
                if (bool.TryParse(enabled, out var enabledBool))
                    boolValue = enabledBool;
            }
            else if (!string.IsNullOrEmpty(value))
            {
                if (bool.TryParse(value, out var valueBool))
                    boolValue = valueBool;
            }

            // Map known layer IDs to properties
            switch (id.ToLowerInvariant())
            {
                case "base":
                case "background":
                    // Base map layer - generally always enabled
                    break;

                case "coastline":
                    if (boolValue.HasValue)
                        settings.ShowCoastlines = boolValue.Value;
                    break;

                case "land":
                case "water":
                    if (boolValue.HasValue && id.ToLowerInvariant() == "water")
                        settings.ShowWaterFeatures = boolValue.Value;
                    break;

                case "admin-0-boundary":
                case "admin-0-countries":
                case "country-borders":
                    if (boolValue.HasValue)
                        settings.ShowCountryBorders = boolValue.Value;
                    break;

                case "admin-1-boundary":
                case "admin-1-states-provinces":
                case "state-borders":
                    if (boolValue.HasValue)
                        settings.ShowStateBorders = boolValue.Value;
                    break;

                case "admin-1-label":
                case "state-name":
                    // State labels
                    break;

                case "urban-area":
                case "place-city":
                    if (boolValue.HasValue)
                        settings.ShowCityNames = boolValue.Value;
                    break;

                case "road":
                case "road-path":
                case "road-motorway-trunk":
                    if (boolValue.HasValue)
                        settings.ShowStreets = boolValue.Value;
                    break;

                default:
                    // Store unknown layers for reference
                    if (!string.IsNullOrEmpty(value))
                        settings.AdditionalLayers[id] = value;
                    else if (boolValue.HasValue)
                        settings.AdditionalLayers[id] = boolValue.Value.ToString();
                    break;
            }
        }

        return settings;
    }

    /// <summary>
    /// Parse the title from layout-options element if present.
    /// Returns null if no title is found or if the title matches the worksheet name/caption.
    /// </summary>
    private string? ParseTitle(XElement worksheet)
    {
        var layoutOptions = worksheet.Element("layout-options");
        if (layoutOptions == null)
            return null;

        var titleElement = layoutOptions.Element("title");
        if (titleElement == null)
            return null;

        var formattedText = titleElement.Element("formatted-text");
        if (formattedText == null)
            return null;

        // Extract text from <run> elements
        var runs = formattedText.Elements("run").ToList();
        if (!runs.Any())
            return null;

        var titleText = string.Join("", runs.Select(r => r.Value)).Trim();

        // Only return if non-empty
        return string.IsNullOrWhiteSpace(titleText) ? null : titleText;
    }

    /// <summary>
    /// Parse table-specific configuration from a worksheet.
    /// Extracts column order, measure names, formatting, and table structure.
    /// </summary>
    private TableConfiguration? ParseTableConfiguration(XDocument document, XElement worksheet, string visualType)
    {
        // Only parse table configuration for table visual types
        if (!visualType.Contains("Table") && visualType != "Crosstab Table")
            return null;

        var table = worksheet.Element("table");
        if (table == null)
            return null;

        var view = table.Element("view");
        if (view == null)
            return null;

        var config = new TableConfiguration
        {
            TableType = visualType
        };

        var rowsValue = table.Element("rows")?.Value ?? string.Empty;
        var colsValue = table.Element("cols")?.Value ?? string.Empty;

        // Check if using Measure Names
        config.UsesMeasureNames = colsValue.Contains("[:Measure Names]") || rowsValue.Contains("[:Measure Names]");

        // Parse row dimensions and extract field references for later use
        List<(string fieldRef, string displayName)> rowDimensions = new();
        if (!string.IsNullOrEmpty(rowsValue) && !rowsValue.Contains("[:Measure Names]"))
        {
            var rowFields = ExtractFieldReferences(rowsValue);
            foreach (var fieldRef in rowFields)
            {
                var displayName = _nameResolver.GetDisplayName(fieldRef);
                config.RowDimensions.Add(displayName);
                rowDimensions.Add((fieldRef, displayName));
            }
        }

        // Parse columns (especially for Measure Names tables)
        if (config.UsesMeasureNames)
        {
            config.Columns = ParseMeasureNamesColumns(document, worksheet, view, rowDimensions);
        }
        else
        {
            // For non-measure-names tables, parse columns from cols
            config.Columns = ParseRegularTableColumns(worksheet, colsValue);
        }

        // Parse table formatting
        config.Formatting = ParseTableFormatting(worksheet);

        // Check for banding
        var tableBandingStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "table");

        if (tableBandingStyle != null)
        {
            var rowBandSize = tableBandingStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "band-size" &&
                                    f.Attribute("scope")?.Value == "rows");
            config.HasRowBanding = rowBandSize != null && rowBandSize.Attribute("value")?.Value != "0";

            var colBandSize = tableBandingStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "band-size" &&
                                    f.Attribute("scope")?.Value == "cols");
            config.HasColumnBanding = colBandSize != null && colBandSize.Attribute("value")?.Value != "0";
        }

        return config;
    }

    /// <summary>
    /// Parse columns from a Measure Names table by extracting the measure order
    /// </summary>
    private List<TableColumn> ParseMeasureNamesColumns(XDocument document, XElement worksheet, XElement view, List<(string fieldRef, string displayName)> rowDimensions)
    {
        var columns = new List<TableColumn>();

        // Add row dimensions as the first columns (with no aggregation)
        int displayOrder = 0;
        foreach (var (fieldRef, displayName) in rowDimensions)
        {
            columns.Add(new TableColumn
            {
                InternalName = fieldRef,
                DisplayName = displayName,
                DisplayOrder = displayOrder++,
                Aggregation = "None",
                NumberFormat = null
            });
        }

        // Find the manual-sort element for [:Measure Names]
        var manualSort = view.Elements("manual-sort")
            .FirstOrDefault(ms => ms.Attribute("column")?.Value?.Contains("[:Measure Names]") == true);

        if (manualSort == null)
            return columns;

        // Extract the datasource name from the manual-sort column attribute
        var manualSortColumn = manualSort.Attribute("column")?.Value;
        string? datasourceName = null;
        if (!string.IsNullOrEmpty(manualSortColumn))
        {
            var match = System.Text.RegularExpressions.Regex.Match(manualSortColumn, @"\[([^\]]+)\]");
            if (match.Success)
            {
                datasourceName = match.Groups[1].Value;
            }
        }

        // Extract member-alias mappings from the datasource
        Dictionary<string, string> memberAliases = new Dictionary<string, string>();
        if (!string.IsNullOrEmpty(datasourceName))
        {
            memberAliases = ExtractMeasureNameAliases(document, datasourceName);
        }

        var dictionary = manualSort.Element("dictionary");
        if (dictionary == null)
            return columns;

        var buckets = dictionary.Elements("bucket").ToList();

        for (int i = 0; i < buckets.Count; i++)
        {
            var bucketValue = buckets[i].Value;

            // Decode HTML entities and strip surrounding quotes
            bucketValue = System.Net.WebUtility.HtmlDecode(bucketValue);
            bucketValue = bucketValue.Trim('"');

            // Extract the field reference from the bucket (e.g., "[federated.xxx].[sum:Field:qk]")
            var columnInstance = FindColumnInstanceByReference(view, bucketValue);
            if (columnInstance == null)
                continue;

            var column = new TableColumn
            {
                InternalName = bucketValue,
                DisplayOrder = displayOrder + i
            };

            // First, try to look up the member-alias for this column-instance
            if (memberAliases.TryGetValue(bucketValue, out var aliasValue))
            {
                column.DisplayName = aliasValue;
            }
            else
            {
                // Fall back to getting the display name from the column definition
                var columnElement = FindColumnDefinition(view, columnInstance);
                if (columnElement != null)
                {
                    column.DisplayName = columnElement.Attribute("caption")?.Value ??
                                        columnElement.Attribute("name")?.Value ??
                                        bucketValue;
                }
                else
                {
                    column.DisplayName = bucketValue;
                }
            }

            // Get number format and aggregation from column definition
            var columnElement2 = FindColumnDefinition(view, columnInstance);
            if (columnElement2 != null)
            {
                column.NumberFormat = columnElement2.Attribute("default-format")?.Value;
                column.Aggregation = ExtractAggregationFromReference(columnInstance);

                // If this is a custom aggregation, extract the calculation name for linking
                if (column.Aggregation == "Custom")
                {
                    column.CalculationName = ExtractCalculationName(columnElement2);
                }
            }
            else
            {
                column.Aggregation = ExtractAggregationFromReference(columnInstance);
            }

            // Check for field-specific format in style rules
            var fieldFormat = worksheet.Descendants("style-rule")
                .Elements("format")
                .FirstOrDefault(f => f.Attribute("field")?.Value == columnInstance);

            if (fieldFormat != null && fieldFormat.Attribute("attr")?.Value == "text-format")
            {
                column.NumberFormat = fieldFormat.Attribute("value")?.Value;
            }

            columns.Add(column);
        }

        return columns;
    }

    /// <summary>
    /// Parse columns from a regular table (not using Measure Names)
    /// </summary>
    private List<TableColumn> ParseRegularTableColumns(XElement worksheet, string colsValue)
    {
        var columns = new List<TableColumn>();

        if (string.IsNullOrEmpty(colsValue))
            return columns;

        var fieldRefs = ExtractFieldReferences(colsValue);
        for (int i = 0; i < fieldRefs.Count; i++)
        {
            var fieldRef = fieldRefs[i];
            var column = new TableColumn
            {
                InternalName = fieldRef,
                DisplayName = _nameResolver.GetDisplayName(fieldRef),
                DisplayOrder = i,
                Aggregation = ExtractAggregationFromReference(fieldRef)
            };

            columns.Add(column);
        }

        return columns;
    }

    /// <summary>
    /// Convert technical format string to friendly format name
    /// </summary>
    private string? GetFriendlyFormatName(string? formatString)
    {
        if (string.IsNullOrEmpty(formatString))
            return null;

        // Currency formats
        if (formatString.Contains("$") || formatString.Contains("c\""))
            return "Currency";

        // Percentage formats
        if (formatString.Contains("%") || formatString.StartsWith("p"))
            return "Percentage";

        // Number formats with thousands separator
        if (formatString.Contains("#,##0"))
            return "Number";

        // Date formats
        if (formatString.Contains("yyyy") || formatString.Contains("MM") || formatString.Contains("dd"))
            return "Date";

        // Default: return the original format string
        return formatString;
    }

    /// <summary>
    /// Extract member-alias mappings from datasources for [:Measure Names] columns.
    /// Returns a dictionary where the key is the full column-instance reference and the value is the friendly alias.
    /// </summary>
    private Dictionary<string, string> ExtractMeasureNameAliases(XDocument document, string datasourceName)
    {
        var aliases = new Dictionary<string, string>();

        // Find the datasource element
        var datasource = document.Descendants("datasource")
            .FirstOrDefault(ds => ds.Attribute("name")?.Value == datasourceName);

        if (datasource == null)
            return aliases;

        // Find the [:Measure Names] column in the datasource
        var measureNamesColumn = datasource.Elements("column")
            .FirstOrDefault(c => c.Attribute("name")?.Value == "[:Measure Names]");

        if (measureNamesColumn == null)
            return aliases;

        // Extract all alias elements
        var aliasElements = measureNamesColumn.Descendants("alias");
        foreach (var aliasElement in aliasElements)
        {
            var key = aliasElement.Attribute("key")?.Value;
            var value = aliasElement.Attribute("value")?.Value;

            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
            {
                // Decode HTML entities and strip quotes from the key
                key = System.Net.WebUtility.HtmlDecode(key);
                key = key.Trim('"');

                aliases[key] = value;
            }
        }

        return aliases;
    }

    /// <summary>
    /// Find a column-instance element by its full reference
    /// </summary>
    private string? FindColumnInstanceByReference(XElement view, string bucketValue)
    {
        // Extract the field portion from bucket value (e.g., "[datasource].[sum:Field:qk]" -> "sum:Field:qk")
        var match = System.Text.RegularExpressions.Regex.Match(bucketValue, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (!match.Success)
            return null;

        var datasource = match.Groups[1].Value;
        var fieldPart = match.Groups[2].Value;

        // Find the column-instance in datasource-dependencies
        var dependencies = view.Elements("datasource-dependencies")
            .FirstOrDefault(dd => dd.Attribute("datasource")?.Value == datasource);

        if (dependencies == null)
            return null;

        var columnInstance = dependencies.Elements("column-instance")
            .FirstOrDefault(ci => ci.Attribute("name")?.Value?.Contains(fieldPart) == true);

        return columnInstance?.Attribute("name")?.Value;
    }

    /// <summary>
    /// Find a column definition element by column-instance name
    /// </summary>
    private XElement? FindColumnDefinition(XElement view, string columnInstanceName)
    {
        // Extract the column reference from column-instance name
        // e.g., "[sum:Calculation_xxx:qk]" -> look for column with name="[Calculation_xxx]"
        var match = System.Text.RegularExpressions.Regex.Match(columnInstanceName, @"\[([^:]+:)?([^\]:]+):?[^\]]*\]");
        if (!match.Success)
            return null;

        var columnName = match.Groups[2].Value;

        // Find in datasource-dependencies
        foreach (var dependencies in view.Elements("datasource-dependencies"))
        {
            var column = dependencies.Elements("column")
                .FirstOrDefault(c => c.Attribute("name")?.Value?.Contains(columnName) == true);

            if (column != null)
                return column;
        }

        return null;
    }

    /// <summary>
    /// Extract aggregation function from a field reference
    /// </summary>
    private string ExtractAggregationFromReference(string fieldRef)
    {
        if (fieldRef.Contains("sum:")) return "SUM";
        if (fieldRef.Contains("avg:")) return "AVG";
        if (fieldRef.Contains("min:")) return "MIN";
        if (fieldRef.Contains("max:")) return "MAX";
        if (fieldRef.Contains("cnt:") || fieldRef.Contains("count:")) return "COUNT";
        if (fieldRef.Contains("usr:")) return "Custom";

        return "None";
    }

    /// <summary>
    /// Extract the calculation name from a column element (for linking purposes)
    /// </summary>
    private string? ExtractCalculationName(XElement columnElement)
    {
        // Get the caption (friendly name) first, fall back to the name attribute
        var calculationName = columnElement.Attribute("caption")?.Value;
        if (!string.IsNullOrEmpty(calculationName))
            return calculationName;

        // Extract from name attribute - format is usually [Calculation_ID]
        var nameAttr = columnElement.Attribute("name")?.Value;
        if (string.IsNullOrEmpty(nameAttr))
            return null;

        // Remove brackets and "Calculation_" prefix if present
        nameAttr = nameAttr.Trim('[', ']');
        if (nameAttr.StartsWith("Calculation_"))
            return null; // Don't return the internal ID, it's not user-friendly

        return nameAttr;
    }

    /// <summary>
    /// Parse table formatting from style rules
    /// </summary>
    private TableFormatting? ParseTableFormatting(XElement worksheet)
    {
        var formatting = new TableFormatting();
        bool hasAnyFormatting = false;

        // Parse header formatting
        var headerStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "header");

        if (headerStyle != null)
        {
            var bgColor = headerStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "background-color");
            if (bgColor != null)
            {
                formatting.HeaderBackgroundColor = bgColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }

            var rowBandColor = headerStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "band-color" &&
                                    f.Attribute("scope")?.Value == "rows");
            if (rowBandColor != null)
            {
                formatting.RowBandingColor = rowBandColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }
        }

        // Parse field-labels formatting (column headers)
        var fieldLabelsStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "field-labels");

        if (fieldLabelsStyle != null)
        {
            var bgColor = fieldLabelsStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "background-color");
            if (bgColor != null && string.IsNullOrEmpty(formatting.HeaderBackgroundColor))
            {
                formatting.HeaderBackgroundColor = bgColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }

            var textColor = fieldLabelsStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "color");
            if (textColor != null)
            {
                formatting.HeaderTextColor = textColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }

            var fontFamily = fieldLabelsStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "font-family");
            if (fontFamily != null)
            {
                formatting.FontFamily = fontFamily.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }
        }

        // Parse label formatting (row headers and cells)
        var labelStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "label");

        if (labelStyle != null)
        {
            var rowAlignment = labelStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "text-align" &&
                                    f.Attribute("scope")?.Value == "rows");
            if (rowAlignment != null)
            {
                formatting.RowHeaderAlignment = rowAlignment.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }

            var rowVerticalAlign = labelStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "vertical-align" &&
                                    f.Attribute("scope")?.Value == "rows");
            if (rowVerticalAlign != null)
            {
                formatting.RowHeaderVerticalAlignment = rowVerticalAlign.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }
        }

        // Parse pane formatting (cell background, banding)
        var paneStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "pane");

        if (paneStyle != null)
        {
            var rowBandColor = paneStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "band-color" &&
                                    f.Attribute("scope")?.Value == "rows");
            if (rowBandColor != null && string.IsNullOrEmpty(formatting.RowBandingColor))
            {
                formatting.RowBandingColor = rowBandColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }

            var colBandColor = paneStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "band-color" &&
                                    f.Attribute("scope")?.Value == "cols");
            if (colBandColor != null)
            {
                formatting.ColumnBandingColor = colBandColor.Attribute("value")?.Value;
                hasAnyFormatting = true;
            }
        }

        // Parse worksheet-level formatting (font size)
        var worksheetStyle = worksheet.Descendants("style-rule")
            .FirstOrDefault(sr => sr.Attribute("element")?.Value == "worksheet");

        if (worksheetStyle != null)
        {
            var fontSize = worksheetStyle.Elements("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "font-size");
            if (fontSize != null && int.TryParse(fontSize.Attribute("value")?.Value, out var size))
            {
                formatting.FontSize = size;
                hasAnyFormatting = true;
            }
        }

        return hasAnyFormatting ? formatting : null;
    }

    /// <summary>
    /// Get the names of worksheets that are marked as hidden in window elements.
    /// Hidden worksheets have: &lt;window class='worksheet' hidden='true' name='WorksheetName'&gt;
    /// </summary>
    private HashSet<string> GetHiddenWorksheetNames(XDocument document)
    {
        var hiddenNames = new HashSet<string>();

        var windowElements = document.Descendants("window")
            .Where(w => w.Attribute("class")?.Value == "worksheet" &&
                        w.Attribute("hidden")?.Value == "true");

        foreach (var window in windowElements)
        {
            var name = window.Attribute("name")?.Value;
            if (!string.IsNullOrEmpty(name))
            {
                hiddenNames.Add(name);
            }
        }

        return hiddenNames;
    }

    /// <summary>
    /// Parse filters from shared-view elements.
    /// Shared-view filters are dashboard-level filters that will be associated with specific dashboards.
    /// </summary>
    public List<Filter> ParseSharedViewFilters(XDocument document)
    {
        var allSharedFilters = new List<Filter>();
        var sharedViews = document.Descendants("shared-view").ToList();

        foreach (var sharedView in sharedViews)
        {
            // Get the datasource name from the shared-view
            var datasourceName = sharedView.Attribute("name")?.Value ?? string.Empty;

            // Parse all filters in this shared-view
            var sharedFilters = ParseFilters(sharedView);

            if (sharedFilters.Any())
            {
                // Add filters to the collection (avoiding duplicates)
                foreach (var filter in sharedFilters)
                {
                    if (!allSharedFilters.Any(f => f.Field == filter.Field && f.Type == filter.Type))
                    {
                        allSharedFilters.Add(filter);
                    }
                }
            }
        }

        return allSharedFilters;
    }
}
