using System.Xml.Linq;
using System.Text.RegularExpressions;
using NeonScribe.Tableau.Core.Models;
using NeonScribe.Tableau.Core.Resolution;

namespace NeonScribe.Tableau.Core.Parsers;

public class DashboardParser
{
    private readonly NameResolver _nameResolver;
    private List<Parameter> _parameters = new();
    private List<DataSource> _dataSources = new();
    private XDocument? _document;

    public DashboardParser(NameResolver nameResolver)
    {
        _nameResolver = nameResolver;
    }

    public List<Dashboard> ParseDashboards(XDocument document, List<Worksheet> worksheets, List<Parameter> parameters, List<DataSource> dataSources)
    {
        _parameters = parameters;
        _dataSources = dataSources;
        _document = document;
        var dashboards = new List<Dashboard>();

        var dbElements = document.Descendants("dashboard").ToList();

        // Build a map of worksheet name -> count of dashboards it appears on (across ALL zone types)
        // This is used to classify worksheets as single-dashboard (→ Visuals) or multi-dashboard (→ Supporting Worksheets)
        var worksheetDashboardCount = BuildWorksheetDashboardCountMap(dbElements, worksheets);

        foreach (var dbElement in dbElements)
        {
            var internalName = dbElement.Attribute("name")?.Value ?? string.Empty;
            var explicitCaption = dbElement.Attribute("caption")?.Value;
            var extractedTitle = ExtractDashboardTitleFromTextZone(dbElement);

            // Dashboard naming priority:
            // 1. Explicit caption attribute (user-defined name)
            // 2. Internal name attribute (often more specific than extracted title)
            // 3. Extracted title from text zone (fallback only)
            var displayCaption = explicitCaption ?? internalName;
            if (string.IsNullOrEmpty(displayCaption))
            {
                displayCaption = extractedTitle ?? "Unnamed Dashboard";
            }

            var dashboard = new Dashboard
            {
                Name = internalName,
                Caption = displayCaption,
                ExtractedTitle = extractedTitle // Store for potential display
            };

            // Parse size
            dashboard.Size = ParseSize(dbElement);

            // Parse color theme
            dashboard.ColorTheme = ParseColorTheme(dbElement, worksheets);

            // Parse zones
            dashboard.Zones = ParseZones(dbElement, displayCaption);

            // Parse actions
            dashboard.Actions = ParseActions(dbElement);

            // Extract visible dashboard filters from filter zones (type-v2='filter')
            // This now includes source field, scope, dynamic filter handling, etc.
            var (dashboardFilters, visualFilters) = ExtractVisibleFiltersWithScope(dbElement, worksheets);
            dashboard.Filters = dashboardFilters;

            // Build visual groups from zone hierarchy (title-worksheet associations)
            // Single-dashboard worksheets without layout-cache are now included in Visuals
            dashboard.VisualGroups = BuildVisualGroups(dbElement, worksheets, dashboard.Filters, visualFilters, worksheetDashboardCount);

            // Build flat visuals list from groups (for backward compatibility)
            dashboard.Visuals = dashboard.VisualGroups.SelectMany(g => g.Visuals).ToList();

            dashboards.Add(dashboard);
        }

        // Mark shared worksheets (worksheets that appear on multiple dashboards)
        MarkSharedWorksheets(dashboards);

        // Build supporting worksheet references (all worksheets without layout-cache)
        BuildSupportingWorksheetReferences(dashboards, worksheets, document);

        return dashboards;
    }

    /// <summary>
    /// Extract a dashboard title from text zones in the dashboard layout.
    /// Looks for prominent text (large font, bold) that likely represents the title.
    /// </summary>
    private string? ExtractDashboardTitleFromTextZone(XElement dashboard)
    {
        var mainZonesElement = dashboard.Element("zones");
        if (mainZonesElement == null)
            return null;

        // Find text zones - looking for formatted text with large font size (likely a title)
        var textZones = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type")?.Value == "text" ||
                        z.Attribute("type-v2")?.Value == "text")
            .ToList();

        string? bestTitle = null;
        int bestFontSize = 0;

        foreach (var textZone in textZones)
        {
            var formattedText = textZone.Element("formatted-text");
            if (formattedText == null)
                continue;

            // Look through all run elements
            foreach (var run in formattedText.Elements("run"))
            {
                var fontSizeStr = run.Attribute("fontsize")?.Value;
                var isBold = run.Attribute("bold")?.Value == "true";
                var text = run.Value?.Trim();

                if (string.IsNullOrEmpty(text))
                    continue;

                // Parse font size
                if (int.TryParse(fontSizeStr, out var fontSize))
                {
                    // Prioritize larger fonts and bold text (typical title characteristics)
                    // Only consider text with fontsize >= 16 as potential titles
                    var score = fontSize + (isBold ? 5 : 0);

                    if (fontSize >= 16 && score > bestFontSize)
                    {
                        bestFontSize = score;
                        // Clean up the title - remove trailing pipes, colons, etc.
                        bestTitle = text.TrimEnd('|', ':', '-', ' ');
                    }
                }
            }
        }

        return bestTitle;
    }

    private DashboardSize ParseSize(XElement dashboard)
    {
        var sizeElement = dashboard.Element("size");

        return new DashboardSize
        {
            Width = int.TryParse(sizeElement?.Attribute("maxwidth")?.Value, out var w) ? w : 1400,
            Height = int.TryParse(sizeElement?.Attribute("maxheight")?.Value, out var h) ? h : 1000
        };
    }

    /// <summary>
    /// Parse color theme information from the dashboard and its worksheets.
    /// Extracts primary color from dashboard format and collects mark colors from visuals.
    /// </summary>
    private DashboardColorTheme? ParseColorTheme(XElement dashboard, List<Worksheet> worksheets)
    {
        var colorTheme = new DashboardColorTheme();
        var hasColorInfo = false;

        // Look for primary color in dashboard's style element
        // This is typically in <style><style-rule element='worksheet'><format attr='color' value='#xxx' />
        var styleElement = dashboard.Element("style");
        if (styleElement != null)
        {
            var colorFormat = styleElement.Descendants("format")
                .FirstOrDefault(f => f.Attribute("attr")?.Value == "color");
            if (colorFormat != null)
            {
                var colorValue = colorFormat.Attribute("value")?.Value;
                if (!string.IsNullOrEmpty(colorValue) && colorValue.StartsWith("#"))
                {
                    colorTheme.PrimaryColor = colorValue;
                    hasColorInfo = true;
                }
            }
        }

        // Collect distinct mark colors from worksheets that appear on this dashboard
        var worksheetNames = dashboard.Element("zones")?.Descendants("zone")
            .Where(z => z.Attribute("name") != null && z.Element("layout-cache") != null)
            .Select(z => z.Attribute("name")?.Value)
            .Where(n => n != null)
            .Distinct()
            .ToHashSet() ?? new HashSet<string?>();

        var markColors = new HashSet<string>();
        string? paletteName = null;

        foreach (var worksheet in worksheets.Where(w => worksheetNames.Contains(w.Name)))
        {
            // Collect mark colors
            if (!string.IsNullOrEmpty(worksheet.MarkEncodings?.Color?.MarkColor))
            {
                markColors.Add(worksheet.MarkEncodings.Color.MarkColor);
            }

            // Collect palette name (use first one found)
            if (paletteName == null && !string.IsNullOrEmpty(worksheet.MarkEncodings?.Color?.PaletteName))
            {
                paletteName = worksheet.MarkEncodings.Color.PaletteName;
            }
        }

        if (markColors.Any())
        {
            colorTheme.MarkColors = markColors.ToList();
            hasColorInfo = true;
        }

        if (!string.IsNullOrEmpty(paletteName))
        {
            colorTheme.PaletteName = paletteName;
            hasColorInfo = true;
        }

        return hasColorInfo ? colorTheme : null;
    }

    private List<DashboardZone> ParseZones(XElement dashboard, string dashboardCaption)
    {
        var zones = new List<DashboardZone>();

        // Only search in the main <zones> element, not in <devicelayouts> which contains
        // duplicate zone definitions for different device layouts (phone, tablet, etc.)
        var mainZonesElement = dashboard.Element("zones");
        if (mainZonesElement == null)
            return zones;

        // Pre-compute dynamic filter controllers so we can resolve their display names
        // (e.g., a CASE-based calculated field swapped by a parameter shows the parameter's default value)
        var dynamicFilterControllers = FindDynamicFilterControllers(dashboard);

        // Build a lookup from zone id -> parent element so we can expand filter zones
        // that are stacked under a text label inside a vertical layout-flow container.
        var zoneElements = mainZonesElement.Descendants("zone").ToList();
        var parentByZoneId = new Dictionary<int, XElement>();
        foreach (var el in mainZonesElement.Descendants("zone"))
        {
            foreach (var child in el.Elements("zone"))
            {
                if (int.TryParse(child.Attribute("id")?.Value, out var childId))
                    parentByZoneId[childId] = el;
            }
        }

        foreach (var zoneElement in zoneElements)
        {
            var zone = new DashboardZone
            {
                Type = zoneElement.Attribute("type")?.Value ?? "layout",
                Id = int.TryParse(zoneElement.Attribute("id")?.Value, out var id) ? id : 0
            };

            // Parse position
            zone.Position = new ZonePosition
            {
                X = int.TryParse(zoneElement.Attribute("x")?.Value, out var x) ? x : 0,
                Y = int.TryParse(zoneElement.Attribute("y")?.Value, out var y) ? y : 0,
                Width = int.TryParse(zoneElement.Attribute("w")?.Value, out var w) ? w : 0,
                Height = int.TryParse(zoneElement.Attribute("h")?.Value, out var h) ? h : 0
            };

            // Parse type-specific attributes
            switch (zone.Type)
            {
                case "text":
                    zone.Content = zoneElement.Element("text")?.Value;
                    break;

                case "worksheet":
                    zone.Worksheet = zoneElement.Attribute("name")?.Value;
                    break;

                case "parameter-control":
                    zone.Parameter = _nameResolver.GetDisplayName(zoneElement.Attribute("param")?.Value ?? string.Empty);
                    break;

                case "bitmap":
                    var style = zoneElement.Element("zone-style");
                    var imgFormat = style?.Elements("format")
                        .FirstOrDefault(f => f.Attribute("attr")?.Value == "background-image");
                    zone.ImagePath = imgFormat?.Attribute("value")?.Value;
                    break;
            }

            // Handle modern Tableau format zones identified by type-v2 or name+layout-cache
            if (zone.Type == "layout")
            {
                var typeV2 = zoneElement.Attribute("type-v2")?.Value;
                var hasName = zoneElement.Attribute("name") != null;
                var hasLayoutCache = zoneElement.Element("layout-cache") != null;
                var hasTypeV2 = typeV2 != null;

                if (typeV2 == "title")
                {
                    zone.Type = "title";
                    // Try to extract custom title text; fall back to dashboard caption (<Sheet Name> placeholder)
                    var formattedText = zoneElement.Element("formatted-text");
                    var customText = formattedText != null
                        ? string.Concat(formattedText.Descendants("run").Select(r => r.Value)).Trim()
                        : null;
                    zone.Content = string.IsNullOrEmpty(customText) ? dashboardCaption : customText;
                }
                else if (typeV2 == "text" || typeV2 == "empty")
                {
                    // Inline text labels (e.g., filter sub-titles) and spacers — skip, not standalone content
                    zone.Type = "skip";
                }
                else if (typeV2 == "filter")
                {
                    zone.Type = "filter";
                    var filterParam = zoneElement.Attribute("param")?.Value ?? string.Empty;
                    // Check if this is a dynamic filter (CASE-based calc field driven by a parameter)
                    var dynamicName = ResolveDynamicFilterName(filterParam, dynamicFilterControllers);
                    zone.Parameter = dynamicName ?? _nameResolver.GetFilterDisplayName(filterParam);

                    // If this filter is inside a vertical layout-flow that also contains a text label,
                    // expand the filter's bounds to fill the parent container (the text label is a title,
                    // not a separate visual element — visually the whole container is one filter control).
                    if (int.TryParse(zoneElement.Attribute("id")?.Value, out var zoneId)
                        && parentByZoneId.TryGetValue(zoneId, out var parentEl)
                        && parentEl.Attribute("type-v2")?.Value == "layout-flow"
                        && parentEl.Attribute("param")?.Value == "vert"
                        && parentEl.Elements("zone").Any(s => s.Attribute("type-v2")?.Value == "text"))
                    {
                        zone.Position = new ZonePosition
                        {
                            X = int.TryParse(parentEl.Attribute("x")?.Value, out var px) ? px : zone.Position.X,
                            Y = int.TryParse(parentEl.Attribute("y")?.Value, out var py) ? py : zone.Position.Y,
                            Width = int.TryParse(parentEl.Attribute("w")?.Value, out var pw) ? pw : zone.Position.Width,
                            Height = int.TryParse(parentEl.Attribute("h")?.Value, out var ph) ? ph : zone.Position.Height,
                        };
                    }
                }
                else if (typeV2 == "color")
                {
                    zone.Type = "color-legend";
                    // Use the param's display name (e.g., "Applications By") rather than the worksheet name
                    var colorParam = zoneElement.Attribute("param")?.Value ?? string.Empty;
                    zone.Parameter = !string.IsNullOrEmpty(colorParam)
                        ? _nameResolver.GetFilterDisplayName(colorParam)
                        : zoneElement.Attribute("name")?.Value;
                }
                else if (typeV2 == "paramctrl")
                {
                    zone.Type = "parameter-control";
                    zone.Parameter = _nameResolver.GetDisplayName(zoneElement.Attribute("param")?.Value ?? string.Empty);
                }
                else if (hasName && hasLayoutCache && !hasTypeV2)
                {
                    // This is a worksheet zone in modern Tableau format (interactive visual)
                    zone.Type = "worksheet";
                    zone.Worksheet = zoneElement.Attribute("name")?.Value;
                }
                else if (hasName && !hasLayoutCache && !hasTypeV2)
                {
                    // This is a background worksheet zone (no layout-cache = non-interactive/decorative)
                    zone.Type = "background-worksheet";
                    zone.Worksheet = zoneElement.Attribute("name")?.Value;
                }
            }

            // Parse style
            zone.Style = ParseZoneStyle(zoneElement);

            zones.Add(zone);
        }

        return zones;
    }

    /// <summary>
    /// Extract visible filter controls with full metadata including scope, source field, and dynamic filter handling.
    /// Returns two lists: dashboard-wide filters and visual-specific filters (keyed by worksheet name).
    ///
    /// Scope determination:
    /// - For filters (type-v2='filter'): The 'name' attribute indicates which worksheet provides the filter values,
    ///   NOT which worksheet it applies to. Most filters apply dashboard-wide. Only filters that are explicitly
    ///   placed near a specific visual (detected by checking if filter is only referenced by one worksheet) are visual-specific.
    /// - For parameters (type-v2='paramctrl'): Check parent zones to see if parameter is grouped with a specific visual.
    /// </summary>
    private (List<Filter> dashboardFilters, Dictionary<string, List<Filter>> visualFilters) ExtractVisibleFiltersWithScope(
        XElement dashboard, List<Worksheet> worksheets)
    {
        var dashboardFilters = new List<Filter>();
        var visualFilters = new Dictionary<string, List<Filter>>();

        var mainZonesElement = dashboard.Element("zones");
        if (mainZonesElement == null)
            return (dashboardFilters, visualFilters);

        // Build a set of all worksheet names on this dashboard for scope detection
        var dashboardWorksheetNames = new HashSet<string>();
        foreach (var zone in mainZonesElement.Descendants("zone"))
        {
            var worksheetName = zone.Attribute("name")?.Value;
            var hasLayoutCache = zone.Element("layout-cache") != null;
            var hasTypeV2 = zone.Attribute("type-v2") != null;
            var isWorksheetType = zone.Attribute("type")?.Value == "worksheet";

            if (!string.IsNullOrEmpty(worksheetName) && (isWorksheetType || (hasLayoutCache && !hasTypeV2)))
            {
                dashboardWorksheetNames.Add(worksheetName);
            }
        }

        // Track which parameters control dynamic filters (for naming dynamic filters)
        var dynamicFilterControllers = FindDynamicFilterControllers(dashboard);

        // Identify visual-specific parameter/filter contexts by finding worksheet zones and their siblings
        var visualSpecificContexts = FindVisualSpecificContexts(mainZonesElement, dashboardWorksheetNames);

        // Find zones with type-v2='filter' OR type='filter' - these are the visible filter controls
        // (older Tableau uses type='filter', newer uses type-v2='filter')
        var filterZones = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type-v2")?.Value == "filter" ||
                        z.Attribute("type")?.Value == "filter")
            .ToList();

        foreach (var filterZone in filterZones)
        {
            var param = filterZone.Attribute("param")?.Value;
            if (string.IsNullOrEmpty(param))
                continue;

            var filter = CreateFilterFromZone(filterZone, param, dynamicFilterControllers, dashboardWorksheetNames);

            // Check if this filter is in a visual-specific context
            var visualContext = GetVisualSpecificContext(filterZone, visualSpecificContexts);
            if (!string.IsNullOrEmpty(visualContext))
            {
                filter.AppliesTo = visualContext;
                if (!visualFilters.ContainsKey(visualContext))
                    visualFilters[visualContext] = new List<Filter>();
                visualFilters[visualContext].Add(filter);
            }
            else
            {
                // Dashboard-wide filter
                dashboardFilters.Add(filter);
            }
        }

        // Process parameter controls
        // (older Tableau uses type='parameter-control', newer uses type-v2='paramctrl')
        var paramZones = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type-v2")?.Value == "paramctrl" ||
                        z.Attribute("type")?.Value == "parameter-control")
            .ToList();

        foreach (var paramZone in paramZones)
        {
            var param = paramZone.Attribute("param")?.Value;
            if (string.IsNullOrEmpty(param))
                continue;

            var filter = CreateParameterFilterFromZone(paramZone, param);

            // Check if this parameter is in a visual-specific context
            var visualContext = GetVisualSpecificContext(paramZone, visualSpecificContexts);
            if (!string.IsNullOrEmpty(visualContext))
            {
                filter.AppliesTo = visualContext;
                if (!visualFilters.ContainsKey(visualContext))
                    visualFilters[visualContext] = new List<Filter>();
                visualFilters[visualContext].Add(filter);
            }
            else
            {
                dashboardFilters.Add(filter);
            }
        }

        // Sort filters by visual position (top-to-bottom, left-to-right)
        var sortedDashboardFilters = SortFiltersByPosition(dashboardFilters);

        // Sort visual-specific filters as well
        var sortedVisualFilters = visualFilters.ToDictionary(
            kvp => kvp.Key,
            kvp => SortFiltersByPosition(kvp.Value)
        );

        return (sortedDashboardFilters, sortedVisualFilters);
    }

    /// <summary>
    /// Find layout zones that contain a single worksheet (these are visual-specific contexts).
    /// Returns a dictionary mapping zone IDs to the worksheet name they contain.
    /// </summary>
    private Dictionary<int, string> FindVisualSpecificContexts(XElement mainZonesElement, HashSet<string> dashboardWorksheetNames)
    {
        var contexts = new Dictionary<int, string>();

        // Find layout-flow zones that contain exactly one worksheet zone
        var layoutFlowZones = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type-v2")?.Value == "layout-flow" ||
                        z.Attribute("type-v2")?.Value == "layout-basic")
            .ToList();

        foreach (var layoutZone in layoutFlowZones)
        {
            var zoneId = int.TryParse(layoutZone.Attribute("id")?.Value, out var id) ? id : 0;
            if (zoneId == 0) continue;

            // Find worksheet zones within this layout
            var worksheetZonesInLayout = layoutZone.Elements("zone")
                .Where(z =>
                {
                    var name = z.Attribute("name")?.Value;
                    var hasLayoutCache = z.Element("layout-cache") != null;
                    var typeV2 = z.Attribute("type-v2")?.Value;
                    return !string.IsNullOrEmpty(name) &&
                           dashboardWorksheetNames.Contains(name) &&
                           (hasLayoutCache || typeV2 == null);
                })
                .Select(z => z.Attribute("name")?.Value)
                .Where(n => n != null)
                .Distinct()
                .ToList();

            // If this layout contains exactly one worksheet, mark it as visual-specific context
            if (worksheetZonesInLayout.Count == 1)
            {
                contexts[zoneId] = worksheetZonesInLayout[0]!;
            }
        }

        return contexts;
    }

    /// <summary>
    /// Check if a zone (filter or parameter) is in a visual-specific context.
    /// Returns the worksheet name if visual-specific, null otherwise.
    /// </summary>
    private string? GetVisualSpecificContext(XElement zone, Dictionary<int, string> visualSpecificContexts)
    {
        // Walk up the parent chain to find if we're in a visual-specific layout
        var parent = zone.Parent;
        while (parent != null)
        {
            var parentId = int.TryParse(parent.Attribute("id")?.Value, out var id) ? id : 0;
            if (parentId > 0 && visualSpecificContexts.TryGetValue(parentId, out var worksheetName))
            {
                return worksheetName;
            }
            parent = parent.Parent;
        }
        return null;
    }

    /// <summary>
    /// Find parameters that control dynamic/calculated filters (CASE statements based on parameter)
    /// Returns a mapping of calculated field internal name -> controlling parameter
    /// </summary>
    private Dictionary<string, Parameter> FindDynamicFilterControllers(XElement dashboard)
    {
        var controllers = new Dictionary<string, Parameter>();

        // Look through all calculated fields to find ones that use CASE [Parameters].[ParameterName]
        foreach (var dataSource in _dataSources)
        {
            foreach (var field in dataSource.Fields.Where(f => f.IsCalculated && !string.IsNullOrEmpty(f.Formula)))
            {
                // Check if formula starts with CASE [Parameters].[...]
                var paramMatch = Regex.Match(field.Formula ?? "", @"CASE\s+\[Parameters\]\.\[([^\]]+)\]", RegexOptions.IgnoreCase);
                if (paramMatch.Success)
                {
                    var paramInternalName = paramMatch.Groups[1].Value;
                    // Find the matching parameter
                    var controllingParam = _parameters.FirstOrDefault(p =>
                        p.InternalName == $"[{paramInternalName}]" ||
                        p.InternalName == paramInternalName ||
                        p.Caption == paramInternalName);

                    if (controllingParam != null)
                    {
                        controllers[field.InternalName] = controllingParam;
                    }
                }
            }
        }

        return controllers;
    }

    /// <summary>
    /// If the filter param refers to a calculated field controlled by a parameter,
    /// returns the parameter's default display name (e.g., "Age Group").
    /// Returns null if not a dynamic filter.
    /// </summary>
    private string? ResolveDynamicFilterName(string param, Dictionary<string, Parameter> dynamicFilterControllers)
    {
        foreach (var kvp in dynamicFilterControllers)
        {
            var calcFieldName = kvp.Key.Trim('[', ']');
            if (param.Contains(calcFieldName))
                return kvp.Value.DefaultDisplayName;
        }
        return null;
    }

    /// <summary>
    /// Create a Filter object from a filter zone with full metadata
    /// </summary>
    private Filter CreateFilterFromZone(XElement filterZone, string param, Dictionary<string, Parameter> dynamicFilterControllers, HashSet<string> worksheetNames)
    {
        // Extract source field from param (e.g., [federated.xxx].[none:FieldName:nk])
        var sourceField = ExtractSourceFieldFromParam(param);

        // Check if this is a dynamic filter controlled by a parameter
        var isDynamicFilter = false;
        Parameter? controllingParameter = null;
        string? dynamicName = null;

        // Check if the source field is a calculated field that's controlled by a parameter
        // The param format is [datasource].[prefix:FieldName:suffix]
        // The InternalName format is [FieldName]
        foreach (var kvp in dynamicFilterControllers)
        {
            // Extract just the field name from InternalName (remove brackets)
            var calcFieldName = kvp.Key.Trim('[', ']');

            // Check if the param contains this calculated field name
            if (param.Contains(calcFieldName) || sourceField == calcFieldName)
            {
                isDynamicFilter = true;
                controllingParameter = kvp.Value;
                // Use the parameter's default display name as the filter name
                dynamicName = controllingParameter.DefaultDisplayName;
                break;
            }
        }

        // Try to get custom title from formatted-text element first
        var customTitle = GetCustomTitle(filterZone);

        // Determine display name
        string displayName;
        if (!string.IsNullOrEmpty(customTitle))
        {
            displayName = customTitle;
        }
        else if (isDynamicFilter && !string.IsNullOrEmpty(dynamicName))
        {
            displayName = dynamicName;
        }
        else
        {
            displayName = _nameResolver.GetFilterDisplayName(param);
        }

        var filter = new Filter
        {
            Field = displayName,
            Type = "categorical",
            FilterType = "Categorical (List)",
            ControlType = MapFilterMode(filterZone.Attribute("mode")?.Value),
            SourceField = _nameResolver.GetFilterDisplayName(param), // The actual underlying field
            InternalParam = param,
            Lineage = _nameResolver.BuildFieldLineage(param), // Build full field lineage
            Position = ExtractZonePosition(filterZone) // Capture position for ordering
        };

        // Add note for dynamic filters
        if (isDynamicFilter && controllingParameter != null)
        {
            filter.Notes = $"Dynamic - controlled by '{controllingParameter.Caption}' parameter";
            // The source field for dynamic filters is the calculated field
            filter.SourceField = "Calculated Field (varies based on parameter selection)";
            // Update lineage to note it's dynamic
            if (filter.Lineage != null)
            {
                filter.Lineage.IsCalculated = true;
                filter.Lineage.Formula = "Dynamic (varies based on parameter selection)";
            }
            // Expose the controlling parameter's allowed values so the card can show the distinct list
            if (controllingParameter.ValueAliases != null && controllingParameter.ValueAliases.Any())
                filter.AllowedValues = controllingParameter.ValueAliases.Values.ToList();
            else if (controllingParameter.AllowableValues?.Type == "list" && controllingParameter.AllowableValues.DisplayNames != null)
                filter.AllowedValues = controllingParameter.AllowableValues.DisplayNames;
        }

        // Extract default selection from the worksheet's filter definitions
        // The 'name' attribute on the filter zone indicates which worksheet provides the filter values
        var worksheetName = filterZone.Attribute("name")?.Value;
        if (!string.IsNullOrEmpty(worksheetName))
        {
            var defaultSelection = ExtractFilterDefaultSelection(worksheetName, param);
            if (!string.IsNullOrEmpty(defaultSelection))
            {
                filter.DefaultSelection = defaultSelection;
            }
        }

        return filter;
    }

    /// <summary>
    /// Extract the default selection value for a filter from the worksheet's filter definitions.
    /// Looks up the worksheet XML and finds the groupfilter member attribute for the matching filter.
    /// If not found in the worksheet, checks shared-views for datasource-level defaults.
    /// </summary>
    /// <param name="worksheetName">The worksheet that provides the filter values</param>
    /// <param name="param">The full param string (e.g., [federated.xxx].[none:Year:ok])</param>
    /// <returns>The default selection value, or null if not found</returns>
    private string? ExtractFilterDefaultSelection(string worksheetName, string param)
    {
        if (_document == null || string.IsNullOrEmpty(worksheetName))
            return null;

        // First, try to find default selection in the worksheet's filter definitions
        var worksheetDefault = ExtractFilterFromWorksheet(worksheetName, param);
        if (worksheetDefault != null)
            return worksheetDefault;

        // If not found in worksheet, check shared-views for datasource-level defaults
        return ExtractFilterFromSharedView(param);
    }

    /// <summary>
    /// Extract default filter selection from a worksheet's filter elements
    /// </summary>
    private string? ExtractFilterFromWorksheet(string worksheetName, string param)
    {
        if (_document == null)
            return null;

        // Find the worksheet element
        var worksheetElement = _document.Descendants("worksheet")
            .FirstOrDefault(w => w.Attribute("name")?.Value == worksheetName);
        if (worksheetElement == null)
            return null;

        // Find filters in the worksheet's view element
        var filters = worksheetElement.Descendants("filter")
            .Where(f => f.Attribute("class")?.Value == "categorical")
            .ToList();

        return ExtractMemberFromFilters(filters, param);
    }

    /// <summary>
    /// Extract default filter selection from shared-view elements (datasource-level defaults)
    /// </summary>
    private string? ExtractFilterFromSharedView(string param)
    {
        if (_document == null)
            return null;

        // Extract datasource name from param
        // param format: [federated.xxx].[none:Year:ok]
        var match = Regex.Match(param, @"^\[([^\]]+)\]");
        if (!match.Success)
            return null;

        var datasourceName = match.Groups[1].Value;

        // Find the shared-view element for this datasource
        var sharedView = _document.Descendants("shared-view")
            .FirstOrDefault(sv => sv.Attribute("name")?.Value == datasourceName);
        if (sharedView == null)
            return null;

        // Find filters in the shared-view
        var filters = sharedView.Descendants("filter")
            .Where(f => f.Attribute("class")?.Value == "categorical")
            .ToList();

        return ExtractMemberFromFilters(filters, param);
    }

    /// <summary>
    /// Extract member value from a list of filter elements
    /// </summary>
    private string? ExtractMemberFromFilters(List<XElement> filters, string param)
    {
        foreach (var filter in filters)
        {
            var column = filter.Attribute("column")?.Value;
            if (string.IsNullOrEmpty(column))
                continue;

            // Check if this filter matches our param
            // param format: [federated.xxx].[none:Year:ok]
            // column format: [federated.xxx].[none:Year:ok]
            if (column == param)
            {
                // Look for groupfilter with function='member' - this is a single selection
                var groupFilter = filter.Element("groupfilter");
                if (groupFilter != null)
                {
                    var function = groupFilter.Attribute("function")?.Value;
                    if (function == "member")
                    {
                        // Single selection - get the member value
                        return groupFilter.Attribute("member")?.Value;
                    }
                    else if (function == "union")
                    {
                        // Multiple selections - collect all member values
                        var members = groupFilter.Descendants("groupfilter")
                            .Where(gf => gf.Attribute("function")?.Value == "member")
                            .Select(gf => gf.Attribute("member")?.Value)
                            .Where(m => m != null)
                            .ToList();
                        if (members.Any())
                        {
                            return string.Join(", ", members);
                        }
                    }
                    else if (function == "level-members")
                    {
                        // "All" selection - check ui-enumeration attribute
                        XNamespace ns = "user";
                        var uiEnum = groupFilter.Attribute(ns + "ui-enumeration")?.Value;
                        if (uiEnum == "all")
                        {
                            return "All";
                        }
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Create a Filter object from a parameter control zone
    /// </summary>
    private Filter CreateParameterFilterFromZone(XElement paramZone, string param)
    {
        // Try to get custom title from formatted-text element first
        var customTitle = GetCustomTitle(paramZone);
        var displayName = !string.IsNullOrEmpty(customTitle)
            ? customTitle
            : _nameResolver.GetDisplayName(param);

        // Find the parameter to get default value info
        var paramName = ExtractParameterName(param);
        var parameter = _parameters.FirstOrDefault(p =>
            p.InternalName == param ||
            p.InternalName == $"[{paramName}]" ||
            p.Caption == displayName);

        // Get the actual parameter name (Caption preferred for display, then Name)
        var actualParamName = parameter?.Caption ?? parameter?.Name ?? displayName;

        var filter = new Filter
        {
            Field = displayName,
            Type = "parameter",
            FilterType = "Parameter",
            ControlType = "Parameter Control",
            SourceField = actualParamName, // The actual parameter name for linking
            InternalParam = param,
            Lineage = new FieldLineage
            {
                DisplayName = displayName,
                InternalName = param,
                BaseFieldName = actualParamName, // Use actual parameter name for linking
                DataSourceName = "Parameters",
                DataType = parameter?.DataType,
                Role = "parameter"
            },
            Position = ExtractZonePosition(paramZone) // Capture position for ordering
        };

        // Set default selection if available
        if (parameter != null && parameter.DefaultDisplayName != null)
        {
            filter.DefaultSelection = parameter.DefaultDisplayName;
        }
        else if (parameter?.DefaultValue != null)
        {
            filter.DefaultSelection = parameter.DefaultValue.ToString();
        }

        // Populate allowed values from parameter's value aliases (distinct list)
        if (parameter?.ValueAliases != null && parameter.ValueAliases.Any())
        {
            filter.AllowedValues = parameter.ValueAliases.Values.ToList();
        }
        else if (parameter?.AllowableValues?.Type == "list" && parameter.AllowableValues.DisplayNames != null)
        {
            filter.AllowedValues = parameter.AllowableValues.DisplayNames;
        }

        return filter;
    }

    /// <summary>
    /// Extract the source field name from a param attribute value
    /// e.g., [federated.xxx].[none:FieldName:nk] -> FieldName
    /// </summary>
    private string ExtractSourceFieldFromParam(string param)
    {
        // Pattern: [datasource].[prefix:FieldName:suffix]
        var match = Regex.Match(param, @"\[([^\]]+)\]\.\[([^\]]+)\]");
        if (match.Success)
        {
            var fieldPart = match.Groups[2].Value;
            // Extract just the field name from prefix:FieldName:suffix
            var colonParts = fieldPart.Split(':');
            if (colonParts.Length >= 2)
            {
                return colonParts[1]; // The middle part is the field name
            }
            return fieldPart;
        }
        return param;
    }

    /// <summary>
    /// Extract parameter name from param attribute
    /// e.g., [Parameters].[Parameter 1] -> Parameter 1
    /// </summary>
    private string ExtractParameterName(string param)
    {
        var match = Regex.Match(param, @"\[Parameters\]\.\[([^\]]+)\]");
        if (match.Success)
        {
            return match.Groups[1].Value;
        }
        // Fallback - try simple bracket extraction
        match = Regex.Match(param, @"\[([^\]]+)\]");
        return match.Success ? match.Groups[1].Value : param;
    }

    /// <summary>
    /// Extract custom title from zone's formatted-text element if present
    /// </summary>
    private string? GetCustomTitle(XElement zone)
    {
        // Check for custom-title attribute first
        if (zone.Attribute("custom-title")?.Value != "true")
            return null;

        // Find formatted-text element and extract run text
        var formattedText = zone.Element("formatted-text");
        if (formattedText == null)
            return null;

        var runElement = formattedText.Element("run");
        return runElement?.Value;
    }

    private string MapFilterMode(string? mode)
    {
        return mode switch
        {
            "checkdropdown" => "Multi-select dropdown",
            "single" => "Single value dropdown",
            "slider" => "Slider",
            "compact" => "Compact list",
            "list" => "List",
            _ => "Dropdown"
        };
    }

    /// <summary>
    /// Extract position from a zone element
    /// </summary>
    private ZonePosition ExtractZonePosition(XElement zone)
    {
        return new ZonePosition
        {
            X = int.TryParse(zone.Attribute("x")?.Value, out var x) ? x : 0,
            Y = int.TryParse(zone.Attribute("y")?.Value, out var y) ? y : 0,
            Width = int.TryParse(zone.Attribute("w")?.Value, out var w) ? w : 0,
            Height = int.TryParse(zone.Attribute("h")?.Value, out var h) ? h : 0
        };
    }

    /// <summary>
    /// Find a title text zone that is positioned directly above a worksheet.
    /// A text zone is considered a title if it overlaps horizontally with the worksheet
    /// and its top edge is above the worksheet's top edge.
    /// </summary>
    private string? FindTitleForWorksheet(ZonePosition worksheetPos, IEnumerable<dynamic> textZones)
    {
        if (worksheetPos == null)
            return null;

        // Find text zones that are:
        // 1. Starts above the worksheet (text Y < worksheet Y) - text zone top is above worksheet top
        // 2. Overlapping horizontally (text zone X range overlaps with worksheet X range)
        // 3. Close enough vertically (within ~10% of dashboard height, or ~10000 units)
        const int maxVerticalDistance = 10000;

        var candidateTitles = new List<(string Text, ZonePosition Position, int Distance)>();

        foreach (var tz in textZones)
        {
            ZonePosition textPos = tz.Position;
            string text = tz.Text;

            if (textPos == null || string.IsNullOrEmpty(text))
                continue;

            // Text zone must start above the worksheet (text top < worksheet top)
            if (textPos.Y >= worksheetPos.Y)
                continue;

            // Check horizontal overlap (or close enough horizontally)
            var wsLeft = worksheetPos.X;
            var wsRight = worksheetPos.X + worksheetPos.Width;
            var textLeft = textPos.X;
            var textRight = textPos.X + textPos.Width;

            // Text zone must overlap horizontally with the worksheet
            var overlaps = !(textRight < wsLeft || textLeft > wsRight);
            if (!overlaps)
                continue;

            // Calculate vertical distance from text zone top to worksheet top
            // This gives us a consistent measure even if zones slightly overlap
            var verticalDistance = worksheetPos.Y - textPos.Y;

            // Skip if too far away (text zone too far above)
            if (verticalDistance < 0 || verticalDistance > maxVerticalDistance)
                continue;

            // Filter out very long text (likely not a title, but a footnote or description)
            // Titles are typically short (< 100 chars)
            if (text.Length > 100)
                continue;

            candidateTitles.Add((text, textPos, verticalDistance));
        }

        // Return the closest text zone (smallest vertical distance - closest to worksheet)
        var bestCandidate = candidateTitles
            .OrderBy(c => c.Distance)
            .FirstOrDefault();

        return bestCandidate.Text;
    }

    /// <summary>
    /// Sort filters by visual position on the dashboard.
    /// Filters are sorted by Y position first (top to bottom), then by X position (left to right).
    /// This ensures vertical filter panels list top-to-bottom, horizontal panels list left-to-right.
    /// </summary>
    private List<Filter> SortFiltersByPosition(List<Filter> filters)
    {
        return filters
            .OrderBy(f => f.Position?.Y ?? 0)
            .ThenBy(f => f.Position?.X ?? 0)
            .ToList();
    }

    private Dictionary<string, object>? ParseZoneStyle(XElement zone)
    {
        var styleElement = zone.Element("zone-style");
        if (styleElement == null)
            return null;

        var style = new Dictionary<string, object>();

        foreach (var format in styleElement.Elements("format"))
        {
            var attr = format.Attribute("attr")?.Value;
            var value = format.Attribute("value")?.Value;

            if (!string.IsNullOrEmpty(attr) && !string.IsNullOrEmpty(value))
            {
                style[ToCamelCase(attr)] = value;
            }
        }

        return style.Any() ? style : null;
    }

    private string ToCamelCase(string text)
    {
        var parts = text.Split('-');
        if (parts.Length == 1)
            return text;

        var result = parts[0];
        for (int i = 1; i < parts.Length; i++)
        {
            if (parts[i].Length > 0)
            {
                result += char.ToUpper(parts[i][0]) + parts[i].Substring(1);
            }
        }

        return result;
    }

    private List<DashboardAction> ParseActions(XElement dashboard)
    {
        var actions = new List<DashboardAction>();

        var actionsElement = dashboard.Element("actions");
        if (actionsElement == null)
            return actions;

        var actionElements = actionsElement.Elements("action").ToList();

        foreach (var actionElement in actionElements)
        {
            var action = new DashboardAction
            {
                Name = actionElement.Attribute("name")?.Value ?? string.Empty,
                Caption = actionElement.Attribute("caption")?.Value ?? string.Empty
            };

            // Parse activation (trigger)
            var activation = actionElement.Element("activation");
            if (activation != null)
            {
                action.AutoClear = activation.Attribute("auto-clear")?.Value == "true";

                var source = activation.Element("source");
                if (source != null)
                {
                    action.Trigger = "select"; // Default
                }
            }

            // Parse source
            var sourceElement = actionElement.Element("source");
            if (sourceElement != null)
            {
                action.Source = new ActionSource
                {
                    Type = sourceElement.Attribute("type")?.Value ?? "worksheet",
                    Name = sourceElement.Attribute("worksheet")?.Value ?? string.Empty
                };
            }

            // Parse target and determine action type
            var targetElement = actionElement.Element("target");
            var urlActionElement = actionElement.Element("url-action");
            var paramActionElement = actionElement.Element("parameter-action");

            if (urlActionElement != null)
            {
                action.Type = "url";
                var urlFormat = urlActionElement.Element("url-format");
                action.UrlFormat = urlFormat?.Value;
                action.TargetBrowser = urlActionElement.Attribute("target")?.Value;
            }
            else if (paramActionElement != null)
            {
                action.Type = "parameter";
                action.SourceField = paramActionElement.Element("source-field")?.Value;
                action.TargetParameter = _nameResolver.GetDisplayName(
                    paramActionElement.Element("target-parameter")?.Value ?? string.Empty);
            }
            else if (targetElement != null)
            {
                // Determine if it's filter or highlight
                var filterElement = targetElement.Element("field");
                action.Type = filterElement != null ? "filter" : "highlight";

                action.Target = new ActionTarget
                {
                    Type = targetElement.Attribute("type")?.Value ?? "worksheet",
                    Name = targetElement.Attribute("worksheet")?.Value ?? string.Empty,
                    FilterNullValues = targetElement.Attribute("filter-null-value")?.Value != "false"
                };

                // Parse fields for filter actions
                if (filterElement != null)
                {
                    action.Fields = new List<string>
                    {
                        _nameResolver.GetDisplayName(filterElement.Value)
                    };
                }
            }

            actions.Add(action);
        }

        return actions;
    }

    /// <summary>
    /// Build visual groups by parsing zone hierarchy to find title-worksheet associations.
    /// Titles are text zones (type-v2='text') that precede worksheet zones within vertical layout containers.
    /// Single-dashboard worksheets without layout-cache are now included in Visuals (not Supporting Worksheets).
    /// </summary>
    private List<VisualGroup> BuildVisualGroups(XElement dashboard, List<Worksheet> worksheets,
        List<Filter> dashboardFilters, Dictionary<string, List<Filter>> visualFilters,
        Dictionary<string, int> worksheetDashboardCount)
    {
        var groups = new List<VisualGroup>();
        var assignedWorksheets = new HashSet<string>();

        var mainZonesElement = dashboard.Element("zones");
        if (mainZonesElement == null)
            return groups;

        // Process vertical layout-flow zones to find title-worksheet groupings
        var verticalContainers = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type-v2")?.Value == "layout-flow"
                        && z.Attribute("param")?.Value == "vert")
            .ToList();

        foreach (var container in verticalContainers)
        {
            var containerGroups = ProcessVerticalContainer(container, worksheets,
                dashboardFilters, visualFilters, assignedWorksheets, worksheetDashboardCount);
            groups.AddRange(containerGroups);
        }

        // Handle orphan worksheets (not in any vertical container with a title)
        // Now includes both layout-cache worksheets AND single-dashboard worksheets without layout-cache
        // Try to associate with nearby text zones based on position
        var allWorksheetZones = mainZonesElement.Descendants("zone")
            .Where(z => IsWorksheetZone(z) || IsSingleDashboardWorksheetZone(z, worksheets, worksheetDashboardCount))
            .ToList();

        // Collect all text zones for position-based title association
        var textZones = mainZonesElement.Descendants("zone")
            .Where(z => z.Attribute("type-v2")?.Value == "text")
            .Select(z => new
            {
                Zone = z,
                Position = ExtractZonePosition(z),
                Text = ExtractTextFromTextZone(z)
            })
            .Where(tz => !string.IsNullOrWhiteSpace(tz.Text))
            .ToList();

        foreach (var wsZone in allWorksheetZones)
        {
            var worksheetName = wsZone.Attribute("name")?.Value;
            if (worksheetName == null || assignedWorksheets.Contains(worksheetName))
                continue;

            var worksheet = worksheets.FirstOrDefault(w => w.Name == worksheetName);
            if (worksheet != null)
            {
                var visual = CreateDashboardVisual(worksheet, wsZone, dashboardFilters, visualFilters);

                // Try to find a title text zone that is directly above this worksheet
                var title = FindTitleForWorksheet(visual.Position, textZones);

                groups.Add(new VisualGroup
                {
                    Title = title,
                    Visuals = new List<DashboardVisual> { visual },
                    Position = visual.Position
                });
                assignedWorksheets.Add(worksheetName);
            }
        }

        // Sort groups by position (Y first, then X)
        return groups.OrderBy(g => g.Position.Y).ThenBy(g => g.Position.X).ToList();
    }

    /// <summary>
    /// Process a vertical layout-flow container to extract title-worksheet groups.
    /// A title applies to all subsequent worksheets until the next title is encountered.
    /// Single-dashboard worksheets without layout-cache are included as visuals.
    /// </summary>
    private List<VisualGroup> ProcessVerticalContainer(XElement container, List<Worksheet> worksheets,
        List<Filter> dashboardFilters, Dictionary<string, List<Filter>> visualFilters,
        HashSet<string> assignedWorksheets, Dictionary<string, int> worksheetDashboardCount)
    {
        var groups = new List<VisualGroup>();
        VisualGroup? currentGroup = null;

        foreach (var child in container.Elements("zone"))
        {
            var typeV2 = child.Attribute("type-v2")?.Value;

            // Text zone - starts a new group
            if (typeV2 == "text")
            {
                // Save current group if it has visuals
                if (currentGroup?.Visuals.Any() == true)
                {
                    groups.Add(currentGroup);
                }

                // Start new group with title from text zone
                var title = ExtractTextFromTextZone(child);
                if (!string.IsNullOrWhiteSpace(title))
                {
                    currentGroup = new VisualGroup
                    {
                        Title = title,
                        Position = ExtractZonePosition(child)
                    };
                }
                else
                {
                    currentGroup = null; // Empty text zones don't start groups
                }
            }
            // Worksheet zone (with layout-cache) OR single-dashboard worksheet without layout-cache - add to current group
            else if (IsWorksheetZone(child) || IsSingleDashboardWorksheetZone(child, worksheets, worksheetDashboardCount))
            {
                var worksheetName = child.Attribute("name")?.Value;
                if (worksheetName != null && !assignedWorksheets.Contains(worksheetName))
                {
                    var worksheet = worksheets.FirstOrDefault(w => w.Name == worksheetName);
                    if (worksheet != null)
                    {
                        if (currentGroup == null)
                        {
                            // No preceding title - create group without title
                            currentGroup = new VisualGroup
                            {
                                Title = null,
                                Position = ExtractZonePosition(child)
                            };
                        }

                        var visual = CreateDashboardVisual(worksheet, child, dashboardFilters, visualFilters);
                        currentGroup.Visuals.Add(visual);
                        assignedWorksheets.Add(worksheetName);
                    }
                }
            }
            // Nested layout container - collect all worksheets within
            else if (typeV2 == "layout-flow" || typeV2 == "layout-basic")
            {
                var nestedWorksheets = CollectWorksheetZonesIncludingSingleDashboard(child, worksheets, worksheetDashboardCount);
                foreach (var wsZone in nestedWorksheets)
                {
                    var worksheetName = wsZone.Attribute("name")?.Value;
                    if (worksheetName != null && !assignedWorksheets.Contains(worksheetName))
                    {
                        var worksheet = worksheets.FirstOrDefault(w => w.Name == worksheetName);
                        if (worksheet != null)
                        {
                            if (currentGroup == null)
                            {
                                currentGroup = new VisualGroup
                                {
                                    Title = null,
                                    Position = ExtractZonePosition(wsZone)
                                };
                            }

                            var visual = CreateDashboardVisual(worksheet, wsZone, dashboardFilters, visualFilters);
                            currentGroup.Visuals.Add(visual);
                            assignedWorksheets.Add(worksheetName);
                        }
                    }
                }
            }
        }

        // Add final group
        if (currentGroup?.Visuals.Any() == true)
        {
            groups.Add(currentGroup);
        }

        return groups;
    }

    /// <summary>
    /// Extract plain text content from a text zone's formatted-text element.
    /// Filters out field references like &lt;[Field]&gt;.
    /// </summary>
    private string? ExtractTextFromTextZone(XElement textZone)
    {
        var formattedText = textZone.Element("formatted-text");
        if (formattedText == null)
            return null;

        var textParts = formattedText.Elements("run")
            .Select(r => r.Value?.Trim())
            .Where(t => !string.IsNullOrEmpty(t) && !t.StartsWith("<[") && !t.StartsWith("&lt;["))
            .ToList();

        var result = string.Join(" ", textParts).Trim();
        return string.IsNullOrEmpty(result) ? null : result;
    }

    /// <summary>
    /// Check if a zone is a worksheet zone (has name attribute and layout-cache, but no type-v2)
    /// </summary>
    private bool IsWorksheetZone(XElement zone)
    {
        var hasName = zone.Attribute("name") != null;
        var hasLayoutCache = zone.Element("layout-cache") != null;
        var typeV2 = zone.Attribute("type-v2")?.Value;

        return hasName && hasLayoutCache && typeV2 == null;
    }

    /// <summary>
    /// Recursively collect all worksheet zones from a zone and its descendants
    /// </summary>
    private List<XElement> CollectWorksheetZones(XElement zone)
    {
        var results = new List<XElement>();

        if (IsWorksheetZone(zone))
        {
            results.Add(zone);
        }

        foreach (var child in zone.Elements("zone"))
        {
            results.AddRange(CollectWorksheetZones(child));
        }

        return results;
    }

    /// <summary>
    /// Recursively collect worksheet zones including single-dashboard worksheets without layout-cache
    /// </summary>
    private List<XElement> CollectWorksheetZonesIncludingSingleDashboard(XElement zone, List<Worksheet> worksheets,
        Dictionary<string, int> worksheetDashboardCount)
    {
        var results = new List<XElement>();

        if (IsWorksheetZone(zone) || IsSingleDashboardWorksheetZone(zone, worksheets, worksheetDashboardCount))
        {
            results.Add(zone);
        }

        foreach (var child in zone.Elements("zone"))
        {
            results.AddRange(CollectWorksheetZonesIncludingSingleDashboard(child, worksheets, worksheetDashboardCount));
        }

        return results;
    }

    /// <summary>
    /// Check if a zone is a worksheet zone without layout-cache that appears on only 1 dashboard.
    /// These are single-dashboard worksheets that should be included in Visuals, not Supporting Worksheets.
    /// </summary>
    private bool IsSingleDashboardWorksheetZone(XElement zone, List<Worksheet> worksheets,
        Dictionary<string, int> worksheetDashboardCount)
    {
        var worksheetName = zone.Attribute("name")?.Value;
        if (string.IsNullOrEmpty(worksheetName))
            return false;

        var hasLayoutCache = zone.Element("layout-cache") != null;
        var typeV2 = zone.Attribute("type-v2")?.Value;

        // Must be a worksheet zone without layout-cache and no type-v2
        if (hasLayoutCache || typeV2 != null)
            return false;

        // Verify it's actually a worksheet
        if (!worksheets.Any(w => w.Name == worksheetName))
            return false;

        // Check if it appears on only 1 dashboard
        if (worksheetDashboardCount.TryGetValue(worksheetName, out var count))
        {
            return count == 1;
        }

        return false;
    }

    /// <summary>
    /// Build a map of worksheet name -> number of dashboards it appears on.
    /// Counts ALL zone references (with or without layout-cache) across all dashboards.
    /// </summary>
    private Dictionary<string, int> BuildWorksheetDashboardCountMap(List<XElement> dbElements, List<Worksheet> worksheets)
    {
        var worksheetNames = worksheets.Select(w => w.Name).ToHashSet();
        var worksheetToDashboards = new Dictionary<string, HashSet<string>>();

        foreach (var dbElement in dbElements)
        {
            var dashboardName = dbElement.Attribute("name")?.Value ?? string.Empty;
            var mainZonesElement = dbElement.Element("zones");
            if (mainZonesElement == null)
                continue;

            // Find all zones that reference worksheets (with or without layout-cache)
            var worksheetZones = mainZonesElement.Descendants("zone")
                .Where(z =>
                {
                    var name = z.Attribute("name")?.Value;
                    var typeV2 = z.Attribute("type-v2")?.Value;
                    // Zone with name, no type-v2, and name matches a worksheet
                    return !string.IsNullOrEmpty(name) && typeV2 == null && worksheetNames.Contains(name);
                })
                .ToList();

            foreach (var zone in worksheetZones)
            {
                var worksheetName = zone.Attribute("name")?.Value!;
                if (!worksheetToDashboards.ContainsKey(worksheetName))
                {
                    worksheetToDashboards[worksheetName] = new HashSet<string>();
                }
                worksheetToDashboards[worksheetName].Add(dashboardName);
            }
        }

        // Convert to count map
        return worksheetToDashboards.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Count);
    }

    /// <summary>
    /// Create a DashboardVisual from a worksheet and its zone element
    /// </summary>
    private DashboardVisual CreateDashboardVisual(Worksheet worksheet, XElement zone,
        List<Filter> dashboardFilters, Dictionary<string, List<Filter>> visualFilters)
    {
        var visual = new DashboardVisual
        {
            Name = worksheet.Name,
            Caption = worksheet.Caption,
            VisualType = worksheet.VisualType,
            MarkType = worksheet.MarkType,
            MarkEncodings = worksheet.MarkEncodings,
            CustomizedLabel = worksheet.CustomizedLabel,
            HasActualTooltipContent = worksheet.HasActualTooltipContent,
            MapConfiguration = worksheet.MapConfiguration,
            Title = worksheet.Title,
            TableConfiguration = worksheet.TableConfiguration,
            WorksheetReference = worksheet.Name,
            Position = ExtractZonePosition(zone),
            FieldsUsed = worksheet.FieldsUsed.Select(f => f.Field).Distinct().ToList(),
            DetailedFieldsUsed = worksheet.FieldsUsed,
            Tooltip = worksheet.Tooltip
        };

        // Dashboard-wide filters apply to this visual
        visual.UsesFilters = dashboardFilters.Select(f => f.Field).ToList();

        // Attach visual-specific filters if any
        if (visualFilters.TryGetValue(worksheet.Name, out var specificFilters))
        {
            visual.VisualSpecificFilters = specificFilters;
        }

        // Generate a basic description based on visual type
        visual.Description = GenerateVisualDescription(worksheet);

        return visual;
    }

    /// <summary>
    /// Build DashboardVisual objects from zones that reference worksheets (legacy flat list)
    /// </summary>
    private List<DashboardVisual> BuildDashboardVisuals(List<DashboardZone> zones, List<Worksheet> worksheets,
        List<Filter> dashboardFilters, Dictionary<string, List<Filter>> visualFilters)
    {
        var visuals = new List<DashboardVisual>();

        // Find all zones that reference worksheets
        var worksheetZones = zones.Where(z => z.Type == "worksheet" && !string.IsNullOrEmpty(z.Worksheet)).ToList();

        foreach (var zone in worksheetZones)
        {
            // Find the corresponding worksheet
            var worksheet = worksheets.FirstOrDefault(w => w.Name == zone.Worksheet);
            if (worksheet == null)
                continue;

            var visual = new DashboardVisual
            {
                Name = worksheet.Name,
                Caption = worksheet.Caption,
                VisualType = worksheet.VisualType,
                MarkType = worksheet.MarkType,
                MarkEncodings = worksheet.MarkEncodings,
                CustomizedLabel = worksheet.CustomizedLabel,
                HasActualTooltipContent = worksheet.HasActualTooltipContent,
                MapConfiguration = worksheet.MapConfiguration,
                Title = worksheet.Title,
                TableConfiguration = worksheet.TableConfiguration,
                WorksheetReference = worksheet.Name,
                Position = zone.Position,
                FieldsUsed = worksheet.FieldsUsed.Select(f => f.Field).Distinct().ToList(),
                DetailedFieldsUsed = worksheet.FieldsUsed,
                Tooltip = worksheet.Tooltip
            };

            // Dashboard-wide filters apply to this visual
            visual.UsesFilters = dashboardFilters.Select(f => f.Field).ToList();

            // Attach visual-specific filters if any
            if (visualFilters.TryGetValue(worksheet.Name, out var specificFilters))
            {
                visual.VisualSpecificFilters = specificFilters;
            }

            // Generate a basic description based on visual type
            visual.Description = GenerateVisualDescription(worksheet);

            visuals.Add(visual);
        }

        return visuals;
    }

    /// <summary>
    /// Generate a verbose description for a visual based on its properties.
    /// For KPIs, explains what each field represents in the display.
    /// For charts, describes the structure including axes, color encoding, and filters.
    /// </summary>
    private string? GenerateVisualDescription(Worksheet worksheet)
    {
        // For KPI/Big Number visuals with customized labels, generate verbose description
        if ((worksheet.VisualType == "KPI / Big Number" ||
             worksheet.VisualType == "KPI Metric") &&
            worksheet.CustomizedLabel?.FieldRoles?.Any() == true)
        {
            return GenerateKpiDescription(worksheet.CustomizedLabel);
        }

        // For Text/Label visuals with customized labels, describe the label content
        if (worksheet.VisualType == "Text / Label" &&
            worksheet.CustomizedLabel?.FieldRoles?.Any() == true)
        {
            return GenerateTextLabelDescription(worksheet);
        }

        // For Text/Label and Text Table visuals without customized labels
        if (worksheet.VisualType == "KPI / Big Number" ||
            worksheet.VisualType == "KPI Metric" ||
            worksheet.VisualType == "Text / Label" ||
            worksheet.VisualType == "Text Table")
        {
            // Get fields from text encodings if available
            var textFields = worksheet.MarkEncodings?.Text?.Fields;
            if (textFields != null && textFields.Any())
            {
                var fieldList = string.Join(", ", textFields.Take(3));
                if (textFields.Count > 3)
                    fieldList += $" +{textFields.Count - 3} more";
                return $"Displays {fieldList}";
            }

            // Fallback to fields used
            var fieldsToShow = worksheet.FieldsUsed.Take(3).Select(f => f.Field).ToList();
            if (fieldsToShow.Any())
            {
                return $"Displays {string.Join(", ", fieldsToShow)}";
            }

            return null; // No description if no fields
        }

        // For chart visuals, generate a more comprehensive description
        return GenerateChartDescription(worksheet);
    }

    /// <summary>
    /// Generate a description for Text/Label visuals with customized labels
    /// </summary>
    private string? GenerateTextLabelDescription(Worksheet worksheet)
    {
        var parts = new List<string>();

        // Describe the content from customized label
        var dynamicFields = worksheet.CustomizedLabel?.FieldRoles?
            .Where(fr => !fr.IsStaticText)
            .Select(fr => fr.FieldName)
            .ToList() ?? new List<string>();

        var staticText = worksheet.CustomizedLabel?.FieldRoles?
            .Where(fr => fr.IsStaticText)
            .Select(fr => fr.FieldName.Trim('"'))
            .ToList() ?? new List<string>();

        if (dynamicFields.Any())
        {
            parts.Add($"Displays: {string.Join(", ", dynamicFields)}");
        }

        if (staticText.Any())
        {
            var combinedStatic = string.Join(" ", staticText).Trim();
            if (!string.IsNullOrWhiteSpace(combinedStatic) && combinedStatic.Length > 0)
            {
                // Truncate if too long
                if (combinedStatic.Length > 50)
                    combinedStatic = combinedStatic.Substring(0, 47) + "...";
                parts.Add($"Label: \"{combinedStatic}\"");
            }
        }

        return parts.Any() ? string.Join(" | ", parts) : null;
    }

    /// <summary>
    /// Generate a description for chart visuals (bar charts, maps, line charts, etc.)
    /// </summary>
    private string? GenerateChartDescription(Worksheet worksheet)
    {
        var descParts = new List<string>();

        // Get the visual type and translate it
        var chartType = TranslateVisualTypeToNaturalLanguage(worksheet.VisualType, worksheet.MarkType);

        // Identify rows/columns (axis fields)
        var rowFields = worksheet.FieldsUsed
            .Where(f => f.Shelf == "Rows")
            .Select(f => f.Field)
            .ToList();
        var columnFields = worksheet.FieldsUsed
            .Where(f => f.Shelf == "Columns")
            .Select(f => f.Field)
            .ToList();

        // Build the description
        if (!string.IsNullOrEmpty(chartType))
        {
            descParts.Add(chartType);
        }

        // Describe what's being shown
        if (rowFields.Any() && columnFields.Any())
        {
            descParts.Add($"showing {string.Join(", ", columnFields)} by {string.Join(", ", rowFields)}");
        }
        else if (rowFields.Any())
        {
            descParts.Add($"by {string.Join(", ", rowFields)}");
        }
        else if (columnFields.Any())
        {
            descParts.Add($"of {string.Join(", ", columnFields)}");
        }

        // Add color encoding info if present
        if (!string.IsNullOrEmpty(worksheet.MarkEncodings?.Color?.Field))
        {
            descParts.Add($"colored by {worksheet.MarkEncodings.Color.Field}");
        }

        // Mention key filters (especially Top N)
        var topNFilter = worksheet.Filters?.FirstOrDefault(f =>
            f.Type == "quantitative" && f.Max != null && f.Min == null &&
            (f.Field.Contains("Rank") || f.Field.Contains("rank") || f.Field.Contains("Top")));
        if (topNFilter != null)
        {
            descParts.Add($"(Top {topNFilter.Max})");
        }

        return descParts.Any() ? string.Join(" ", descParts) : null;
    }

    /// <summary>
    /// Translate visual type codes to natural language
    /// </summary>
    private string TranslateVisualTypeToNaturalLanguage(string? visualType, string? markType)
    {
        // Check mark type first for specific chart types
        var mark = markType?.ToLowerInvariant();
        switch (mark)
        {
            case "bar":
                return "Bar chart";
            case "line":
                return "Line chart";
            case "area":
                return "Area chart";
            case "circle":
            case "shape":
                return "Scatter plot";
            case "square":
                return "Heat map";
            case "polygon":
            case "map":
                return "Map";
            case "pie":
                return "Pie chart";
            case "gantt":
                return "Gantt chart";
        }

        // Fall back to visual type
        return visualType switch
        {
            "Geographic Map" => "Map",
            "Horizontal Bar" or "Vertical Bar" => "Bar chart",
            "Line Chart" => "Line chart",
            "Area Chart" => "Area chart",
            "Scatter Plot" => "Scatter plot",
            "Text / Label" => "Text display",
            "Text Table" => "Table",
            _ => string.IsNullOrEmpty(visualType) ? "Chart" : visualType
        };
    }

    /// <summary>
    /// Generate a verbose description for KPI visuals explaining each field's role
    /// </summary>
    private string GenerateKpiDescription(CustomizedLabel customizedLabel)
    {
        var parts = new List<string>();

        foreach (var fieldRole in customizedLabel.FieldRoles)
        {
            parts.Add($"• {fieldRole.FieldName}: {fieldRole.Role}");
        }

        return string.Join("\n", parts);
    }

    /// <summary>
    /// Mark worksheets that appear on multiple dashboards as shared and populate SharedWorksheets lists.
    /// Also detects worksheets on dashboards without layout-cache (background worksheets that may be shared).
    /// </summary>
    private void MarkSharedWorksheets(List<Dashboard> dashboards)
    {
        // Build a dictionary of worksheet name -> count of appearances in Visuals (layout-cache worksheets)
        var worksheetUsageCount = new Dictionary<string, int>();

        foreach (var dashboard in dashboards)
        {
            foreach (var visual in dashboard.Visuals)
            {
                if (!worksheetUsageCount.ContainsKey(visual.WorksheetReference))
                {
                    worksheetUsageCount[visual.WorksheetReference] = 0;
                }
                worksheetUsageCount[visual.WorksheetReference]++;
            }
        }

        // Mark visuals as shared if they appear on multiple dashboards
        foreach (var dashboard in dashboards)
        {
            foreach (var visual in dashboard.Visuals)
            {
                visual.IsSharedAcrossDashboards = worksheetUsageCount[visual.WorksheetReference] > 1;
            }
        }
    }

    /// <summary>
    /// Build SupportingWorksheets lists for each dashboard.
    /// Only includes worksheets that appear on 2+ dashboards (shared worksheets).
    /// Single-dashboard worksheets without layout-cache are now included in Visuals instead.
    /// </summary>
    public void BuildSupportingWorksheetReferences(List<Dashboard> dashboards, List<Worksheet> worksheets, XDocument document)
    {
        // Build a dictionary tracking which dashboards each worksheet appears on (via zones without layout-cache)
        var worksheetToDashboards = new Dictionary<string, List<string>>();

        var dbElements = document.Descendants("dashboard").ToList();

        foreach (var dbElement in dbElements)
        {
            var dashboardName = dbElement.Attribute("name")?.Value ?? string.Empty;
            var dashboardCaption = dashboards.FirstOrDefault(d => d.Name == dashboardName)?.Caption ?? dashboardName;

            var mainZonesElement = dbElement.Element("zones");
            if (mainZonesElement == null)
                continue;

            // Find all zones that reference worksheets but don't have layout-cache (background/shared worksheets)
            var supportingWorksheetZones = mainZonesElement.Descendants("zone")
                .Where(z =>
                {
                    var hasName = z.Attribute("name") != null;
                    var hasLayoutCache = z.Element("layout-cache") != null;
                    var typeV2 = z.Attribute("type-v2")?.Value;
                    // Zone with name, no layout-cache, no type-v2 = supporting worksheet
                    return hasName && !hasLayoutCache && typeV2 == null;
                })
                .ToList();

            foreach (var zone in supportingWorksheetZones)
            {
                var worksheetName = zone.Attribute("name")?.Value;
                if (string.IsNullOrEmpty(worksheetName))
                    continue;

                // Verify this is actually a worksheet
                if (!worksheets.Any(w => w.Name == worksheetName))
                    continue;

                if (!worksheetToDashboards.ContainsKey(worksheetName))
                {
                    worksheetToDashboards[worksheetName] = new List<string>();
                }

                if (!worksheetToDashboards[worksheetName].Contains(dashboardCaption))
                {
                    worksheetToDashboards[worksheetName].Add(dashboardCaption);
                }
            }
        }

        // Now populate SupportingWorksheets for each dashboard
        // ONLY include worksheets that appear on 2+ dashboards (shared worksheets)
        // Single-dashboard worksheets are now in Visuals
        foreach (var dbElement in dbElements)
        {
            var dashboardName = dbElement.Attribute("name")?.Value ?? string.Empty;
            var dashboard = dashboards.FirstOrDefault(d => d.Name == dashboardName);
            if (dashboard == null)
                continue;

            var mainZonesElement = dbElement.Element("zones");
            if (mainZonesElement == null)
                continue;

            // Find supporting worksheet zones for this dashboard
            var supportingWorksheetZones = mainZonesElement.Descendants("zone")
                .Where(z =>
                {
                    var hasName = z.Attribute("name") != null;
                    var hasLayoutCache = z.Element("layout-cache") != null;
                    var typeV2 = z.Attribute("type-v2")?.Value;
                    return hasName && !hasLayoutCache && typeV2 == null;
                })
                .ToList();

            var addedWorksheets = new HashSet<string>();

            foreach (var zone in supportingWorksheetZones)
            {
                var worksheetName = zone.Attribute("name")?.Value;
                if (string.IsNullOrEmpty(worksheetName) || addedWorksheets.Contains(worksheetName))
                    continue;

                // Find the worksheet
                var worksheet = worksheets.FirstOrDefault(w => w.Name == worksheetName);
                if (worksheet == null)
                    continue;

                // Only include worksheets that appear on 2+ dashboards (shared worksheets)
                // Single-dashboard worksheets are now included in Visuals instead
                var usedIn = worksheetToDashboards.TryGetValue(worksheetName, out var dashboardList) ? dashboardList : new List<string>();
                if (usedIn.Count < 2)
                    continue; // Skip single-dashboard worksheets - they're in Visuals now

                var supportingRef = new SupportingWorksheetReference
                {
                    Name = worksheet.Name,
                    Caption = !string.IsNullOrEmpty(worksheet.Caption) ? worksheet.Caption : worksheet.Name,
                    VisualType = worksheet.VisualType ?? "Unknown",
                    Description = GenerateVisualDescription(worksheet),
                    IsShared = true, // Always true now since we only include shared worksheets
                    UsedInDashboards = usedIn
                };

                dashboard.SupportingWorksheets.Add(supportingRef);
                addedWorksheets.Add(worksheetName);
            }
        }
    }
}
