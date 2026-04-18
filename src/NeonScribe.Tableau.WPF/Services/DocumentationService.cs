using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using NeonScribe.Tableau.Core.Models;
using NeonScribe.Tableau.Core.Parsers;
using NeonScribe.Tableau.Documentation.Generators;

namespace NeonScribe.Tableau.WPF.Services;

public class DocumentationService
{
    public async Task<TableauWorkbook> ParseWorkbookAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            var parser = new WorkbookParser();
            return parser.Parse(filePath);
        });
    }

    public async Task GenerateHtmlAsync(TableauWorkbook workbook, string outputPath)
    {
        await Task.Run(() =>
        {
            var generator = new HtmlGenerator(workbook);
            var html = generator.Generate();
            File.WriteAllText(outputPath, html);
        });
    }

    public async Task GenerateJsonAsync(TableauWorkbook workbook, string outputPath)
    {
        await Task.Run(() =>
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var json = JsonSerializer.Serialize(workbook, options);
            File.WriteAllText(outputPath, json);
        });
    }
}
