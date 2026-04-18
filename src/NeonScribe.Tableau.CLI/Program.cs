using System.Text.Json;
using System.Text.Json.Serialization;
using NeonScribe.Tableau.Core.Parsers;
using NeonScribe.Tableau.Documentation.Generators;

namespace NeonScribe.Tableau.CLI;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Neon Scribe for Tableau (CLI)");
                Console.WriteLine("Usage: NeonScribe.Tableau.CLI <input-file.twb|twbx> [-o output-file] [-f format]");
                Console.WriteLine();
                Console.WriteLine("Arguments:");
                Console.WriteLine("  <input-file>    Path to TWB or TWBX file");
                Console.WriteLine("  -o <output>     Output file path (optional, defaults to stdout)");
                Console.WriteLine("  -f <format>     Output format: json or html (default: html)");
                Console.WriteLine();
                Console.WriteLine("Examples:");
                Console.WriteLine("  NeonScribe.Tableau.CLI workbook.twb -o output.html");
                Console.WriteLine("  NeonScribe.Tableau.CLI workbook.twb -o output.json -f json");
                return;
            }

            string inputFile = args[0];
            string? outputFile = null;
            string format = "html";

            // Parse arguments
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "-o" && i + 1 < args.Length)
                {
                    outputFile = args[i + 1];
                    i++;
                }
                else if (args[i] == "-f" && i + 1 < args.Length)
                {
                    format = args[i + 1].ToLower();
                    i++;
                }
            }

            if (!File.Exists(inputFile))
            {
                Console.Error.WriteLine($"Error: File not found: {inputFile}");
                Environment.Exit(1);
            }

            Console.WriteLine($"Parsing: {inputFile}");

            // Parse the workbook
            var parser = new WorkbookParser();
            var workbook = parser.Parse(inputFile);

            Console.WriteLine($"Successfully parsed workbook:");
            Console.WriteLine($"  Version: {workbook.Version}");
            Console.WriteLine($"  Data Sources: {workbook.DataSourcesCount}");
            Console.WriteLine($"  Parameters: {workbook.ParametersCount}");
            Console.WriteLine($"  Worksheets: {workbook.WorksheetsCount}");
            Console.WriteLine($"  Dashboards: {workbook.DashboardsCount}");
            Console.WriteLine($"  Total Fields: {workbook.TotalFields}");
            Console.WriteLine($"  Calculated Fields: {workbook.CalculatedFields}");

            string output;

            // Generate output based on format
            if (format == "html")
            {
                Console.WriteLine("Generating HTML documentation...");
                var htmlGenerator = new HtmlGenerator(workbook);
                output = htmlGenerator.Generate();
            }
            else if (format == "json")
            {
                // Serialize to JSON
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };
                output = JsonSerializer.Serialize(workbook, options);
            }
            else
            {
                Console.Error.WriteLine($"Error: Unknown format '{format}'. Supported formats: json, html");
                Environment.Exit(1);
                return;
            }

            // Output to file or stdout
            if (!string.IsNullOrEmpty(outputFile))
            {
                File.WriteAllText(outputFile, output);
                Console.WriteLine($"Output written to: {outputFile}");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine($"{format.ToUpper()} Output:");
                Console.WriteLine(output);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            Console.Error.WriteLine(ex.StackTrace);
            Environment.Exit(1);
        }
    }
}
