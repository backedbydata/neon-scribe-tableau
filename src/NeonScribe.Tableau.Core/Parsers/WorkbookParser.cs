using System.Xml.Linq;
using NeonScribe.Tableau.Core.Models;
using NeonScribe.Tableau.Core.Resolution;
using NeonScribe.Tableau.Core.Utilities;

namespace NeonScribe.Tableau.Core.Parsers;

public class WorkbookParser
{
    private NameResolver? _nameResolver;
    private XDocument? _document;

    public TableauWorkbook Parse(string filePath)
    {
        // Handle both TWB and TWBX files
        var twbPath = TwbxExtractor.GetTwbPath(filePath);

        // Load the XML document
        _document = XDocument.Load(twbPath);

        // Build name resolver first (critical for user-friendly names)
        _nameResolver = new NameResolver(_document);

        // Create workbook object
        var workbook = new TableauWorkbook
        {
            FileName = Path.GetFileName(filePath),
            Version = _document.Root?.Attribute("version")?.Value ?? string.Empty,
            SourceBuild = _document.Root?.Attribute("source-build")?.Value ?? string.Empty
        };

        // Parse datasources
        var dataSourceParser = new DataSourceParser(_nameResolver);
        workbook.DataSources = dataSourceParser.ParseDataSources(_document);
        workbook.Parameters = dataSourceParser.ParseParameters(_document);

        // Parse worksheets
        var worksheetParser = new WorksheetParser(_nameResolver);
        workbook.Worksheets = worksheetParser.ParseWorksheets(_document);

        // Parse dashboards - filters are extracted from visible filter zones on each dashboard
        // Pass parameters so we can resolve dynamic filter names and defaults
        var dashboardParser = new DashboardParser(_nameResolver);
        workbook.Dashboards = dashboardParser.ParseDashboards(_document, workbook.Worksheets, workbook.Parameters, workbook.DataSources);

        // Calculate counts and statistics
        workbook.DataSourcesCount = workbook.DataSources.Count;
        workbook.ParametersCount = workbook.Parameters.Count;
        workbook.WorksheetsCount = workbook.Worksheets.Count;
        workbook.DashboardsCount = workbook.Dashboards.Count;
        workbook.StoriesCount = _document.Descendants("story").Count();

        workbook.TotalFields = workbook.DataSources.Sum(ds => ds.Fields.Count);
        workbook.CalculatedFields = workbook.DataSources.Sum(ds => ds.Fields.Count(f => f.IsCalculated));
        workbook.LodCalculations = workbook.DataSources.Sum(ds => ds.Fields.Count(f => f.CalculationType == "LOD"));
        workbook.TableCalculations = workbook.DataSources.Sum(ds => ds.Fields.Count(f => f.CalculationType == "Table Calculation"));
        // Count both dashboard-level filters and worksheet-specific filters
        workbook.Filters = workbook.Dashboards.Sum(db => db.Filters.Count) + workbook.Worksheets.Sum(ws => ws.Filters.Count);
        workbook.Actions = workbook.Dashboards.Sum(db => db.Actions.Count);

        // Build field usage map and calculation dependencies
        BuildFieldUsageMap(workbook);
        BuildCalculationDependencies(workbook);

        // Store name mappings for reference
        workbook.InternalNameToDisplayName = _nameResolver.GetAllMappings();

        return workbook;
    }

    private void BuildFieldUsageMap(TableauWorkbook workbook)
    {
        var fieldUsageMap = new Dictionary<string, FieldUsage>();

        // Iterate through all fields across all data sources
        foreach (var dataSource in workbook.DataSources)
        {
            foreach (var field in dataSource.Fields)
            {
                var usage = new FieldUsage
                {
                    DataSource = dataSource.Caption
                };

                // Find worksheets using this field
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.FieldsUsed.Any(fu => fu.Field == field.Name))
                    {
                        usage.UsedInWorksheets.Add(worksheet.Caption);
                    }

                    if (worksheet.Filters.Any(f => f.Field == field.Name))
                    {
                        usage.UsedInFilters.Add(worksheet.Caption);
                    }
                }

                // Find calculations using this field
                foreach (var ds in workbook.DataSources)
                {
                    foreach (var calcField in ds.Fields.Where(f => f.IsCalculated))
                    {
                        if (calcField.Dependencies?.Contains(field.InternalName) == true ||
                            calcField.Dependencies?.Contains(field.Name) == true)
                        {
                            usage.UsedInCalculations.Add(calcField.Name);
                        }
                    }
                }

                // Find actions using this field
                foreach (var dashboard in workbook.Dashboards)
                {
                    foreach (var action in dashboard.Actions)
                    {
                        if (action.Fields?.Contains(field.Name) == true ||
                            action.FieldsUsed?.Contains(field.Name) == true)
                        {
                            usage.UsedInActions.Add(action.Caption);
                        }
                    }
                }

                fieldUsageMap[field.Name] = usage;
            }
        }

        workbook.FieldUsageMap = fieldUsageMap;
    }

    private void BuildCalculationDependencies(TableauWorkbook workbook)
    {
        var calcDependencies = new Dictionary<string, CalculationDependency>();

        // Get all calculated fields
        var allCalculatedFields = workbook.DataSources
            .SelectMany(ds => ds.Fields.Where(f => f.IsCalculated))
            .ToList();

        foreach (var calcField in allCalculatedFields)
        {
            var dependency = new CalculationDependency
            {
                Formula = calcField.Formula ?? string.Empty,
                Dependencies = calcField.Dependencies ?? new List<string>(),
                Explanation = calcField.Explanation,
                CalculationType = calcField.CalculationType,
                LodType = calcField.LodType
            };

            // Find which other calculations use this one
            foreach (var otherCalc in allCalculatedFields)
            {
                if (otherCalc.Name != calcField.Name &&
                    (otherCalc.Dependencies?.Contains(calcField.InternalName) == true ||
                     otherCalc.Dependencies?.Contains(calcField.Name) == true))
                {
                    dependency.UsedBy.Add(otherCalc.Name);
                }
            }

            // Resolve dependency names to display names
            dependency.Dependencies = dependency.Dependencies
                .Select(d => _nameResolver?.GetDisplayName(d) ?? d)
                .ToList();

            calcDependencies[calcField.Name] = dependency;
        }

        workbook.CalculationDependencies = calcDependencies;
    }
}
