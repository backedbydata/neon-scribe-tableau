using System.Text.RegularExpressions;
using NeonScribe.Tableau.Core.Models;

namespace NeonScribe.Tableau.Core.Utilities;

public class CalculationExplainer
{
    /// <summary>
    /// Parse LOD calculation formula to extract dimensions in scope
    /// </summary>
    public static List<string> ExtractLodDimensions(string formula, string lodType)
    {
        var dimensions = new List<string>();

        if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(lodType))
            return dimensions;

        // Pattern to match LOD scope: { FIXED [Dim1], [Dim2] : aggregation }
        var pattern = $@"{{\s*{lodType}\s+(.*?)\s*:";
        var match = Regex.Match(formula, pattern, RegexOptions.IgnoreCase);

        if (match.Success && match.Groups.Count > 1)
        {
            var dimensionsPart = match.Groups[1].Value;

            // Extract field references [FieldName]
            var fieldPattern = @"\[([^\]]+)\]";
            var fieldMatches = Regex.Matches(dimensionsPart, fieldPattern);

            foreach (Match fieldMatch in fieldMatches)
            {
                if (fieldMatch.Groups.Count > 1)
                {
                    dimensions.Add(fieldMatch.Groups[1].Value);
                }
            }
        }

        return dimensions;
    }

    /// <summary>
    /// Generate natural language explanation for LOD calculation
    /// </summary>
    public static string ExplainLodCalculation(Field field)
    {
        if (string.IsNullOrEmpty(field.LodType) || string.IsNullOrEmpty(field.Formula))
            return string.Empty;

        var aggregation = ExtractAggregationFunction(field.Formula);
        var measureField = ExtractMeasureField(field.Formula);

        return field.LodType switch
        {
            "FIXED" => GenerateFixedExplanation(field, aggregation, measureField),
            "INCLUDE" => GenerateIncludeExplanation(field, aggregation, measureField),
            "EXCLUDE" => GenerateExcludeExplanation(field, aggregation, measureField),
            _ => string.Empty
        };
    }

    /// <summary>
    /// Generate natural language explanation for table calculation
    /// </summary>
    public static string ExplainTableCalculation(Field field)
    {
        if (string.IsNullOrEmpty(field.TableCalcFunction) || string.IsNullOrEmpty(field.Formula))
            return string.Empty;

        return field.TableCalcFunction switch
        {
            "RUNNING_SUM" => $"Calculates a running total (cumulative sum) across the table. Each row shows the sum of all values up to and including that row.",
            "RUNNING_AVG" => $"Calculates a running average across the table. Each row shows the average of all values up to and including that row.",
            "RUNNING_MAX" => $"Calculates a running maximum across the table. Each row shows the maximum value up to and including that row.",
            "RUNNING_MIN" => $"Calculates a running minimum across the table. Each row shows the minimum value up to and including that row.",
            "WINDOW_SUM" => $"Calculates the sum within a moving window of rows in the table.",
            "WINDOW_AVG" => $"Calculates the average within a moving window of rows in the table.",
            "WINDOW_MAX" => $"Calculates the maximum within a moving window of rows in the table.",
            "WINDOW_MIN" => $"Calculates the minimum within a moving window of rows in the table.",
            "LOOKUP" => $"Retrieves values from other rows in the table, allowing you to reference values from previous or next rows.",
            "INDEX" => $"Returns the index (position) of the current row within the partition, starting at 1.",
            "RANK" => $"Assigns a rank to each row within the partition based on the values.",
            "PERCENT_OF_TOTAL" => $"Calculates what percentage each value represents of the total.",
            "PERCENTILE" => $"Calculates the percentile rank of values within the partition.",
            _ => $"Table calculation that computes values based on the data in the view."
        };
    }

    /// <summary>
    /// Detect and return table calculation function name
    /// </summary>
    public static string? DetectTableCalcFunction(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return null;

        var functions = new[]
        {
            "RUNNING_SUM", "RUNNING_AVG", "RUNNING_MAX", "RUNNING_MIN",
            "WINDOW_SUM", "WINDOW_AVG", "WINDOW_MAX", "WINDOW_MIN",
            "LOOKUP", "INDEX", "RANK", "PERCENT_OF_TOTAL", "PERCENTILE"
        };

        foreach (var function in functions)
        {
            if (formula.Contains(function, StringComparison.OrdinalIgnoreCase))
            {
                return function;
            }
        }

        return null;
    }

    private static string GenerateFixedExplanation(Field field, string aggregation, string measureField)
    {
        if (field.LodDimensions == null || !field.LodDimensions.Any())
        {
            return $"Calculates the {aggregation} of {measureField} at the data source level, ignoring all filters and dimensions in the view.";
        }

        var dimList = FormatDimensionList(field.LodDimensions);

        if (field.LodDimensions.Count == 1)
        {
            return $"Calculates the {aggregation} of {measureField} for each {field.LodDimensions[0]}, regardless of other dimensions in the view. This value remains constant for each {field.LodDimensions[0]} even when the view is broken down by other dimensions.";
        }
        else
        {
            return $"Calculates the {aggregation} of {measureField} at the level of {dimList}, regardless of other dimensions in the view. This value is fixed at this specific combination of dimensions.";
        }
    }

    private static string GenerateIncludeExplanation(Field field, string aggregation, string measureField)
    {
        if (field.LodDimensions == null || !field.LodDimensions.Any())
        {
            return $"Calculates the {aggregation} of {measureField} including additional level of detail.";
        }

        var dimList = FormatDimensionList(field.LodDimensions);

        return $"Calculates the {aggregation} of {measureField} including {dimList} in addition to the dimensions already in the view. This adds finer granularity to the calculation beyond what's visible in the view.";
    }

    private static string GenerateExcludeExplanation(Field field, string aggregation, string measureField)
    {
        if (field.LodDimensions == null || !field.LodDimensions.Any())
        {
            return $"Calculates the {aggregation} of {measureField} excluding certain dimensions.";
        }

        var dimList = FormatDimensionList(field.LodDimensions);

        if (field.LodDimensions.Count == 1)
        {
            return $"Calculates the {aggregation} of {measureField} while ignoring {field.LodDimensions[0]}, even if it appears in the view. This aggregates data to a higher level than what's shown.";
        }
        else
        {
            return $"Calculates the {aggregation} of {measureField} while excluding {dimList} from the calculation, even if they appear in the view. This aggregates data to a higher level.";
        }
    }

    private static string ExtractAggregationFunction(string formula)
    {
        var aggPattern = @"\b(SUM|AVG|COUNT|MIN|MAX|MEDIAN|STDEV|VAR)\b";
        var match = Regex.Match(formula, aggPattern, RegexOptions.IgnoreCase);

        if (match.Success)
        {
            return match.Value.ToLower() switch
            {
                "sum" => "sum",
                "avg" => "average",
                "count" => "count",
                "min" => "minimum",
                "max" => "maximum",
                "median" => "median",
                "stdev" => "standard deviation",
                "var" => "variance",
                _ => match.Value.ToLower()
            };
        }

        return "aggregation";
    }

    private static string ExtractMeasureField(string formula)
    {
        // Extract the field name being aggregated, typically in the format SUM([FieldName])
        var pattern = @"(?:SUM|AVG|COUNT|MIN|MAX|MEDIAN|STDEV|VAR)\s*\(\s*\[([^\]]+)\]\s*\)";
        var match = Regex.Match(formula, pattern, RegexOptions.IgnoreCase);

        if (match.Success && match.Groups.Count > 1)
        {
            return match.Groups[1].Value;
        }

        return "the measure";
    }

    private static string FormatDimensionList(List<string> dimensions)
    {
        if (dimensions.Count == 0) return string.Empty;
        if (dimensions.Count == 1) return dimensions[0];
        if (dimensions.Count == 2) return $"{dimensions[0]} and {dimensions[1]}";

        var allButLast = string.Join(", ", dimensions.Take(dimensions.Count - 1));
        return $"{allButLast}, and {dimensions.Last()}";
    }
}
