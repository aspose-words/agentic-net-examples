using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();

    // Generates a styled HTML table from the Items collection.
    public string HtmlTable => GenerateHtmlTable();

    private string GenerateHtmlTable()
    {
        var sb = new StringBuilder();
        sb.Append("<table style='border:1px solid black; border-collapse:collapse;'>");
        sb.Append("<tr>");
        sb.Append("<th style='border:1px solid black; background:#D3D3D3;'>Index</th>");
        sb.Append("<th style='border:1px solid black; background:#D3D3D3;'>Name</th>");
        sb.Append("</tr>");

        foreach (var item in Items)
        {
            sb.Append("<tr>");
            sb.Append($"<td style='border:1px solid black;'>{item.Index}</td>");
            sb.Append($"<td style='border:1px solid black;'>{WebUtility.HtmlEncode(item.Name)}</td>");
            sb.Append("</tr>");
        }

        sb.Append("</table>");
        return sb.ToString();
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // 1. Create sample JSON data file.
        string jsonPath = Path.Combine(workDir, "data.json");
        var sampleData = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON into the model object.
        var model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath))!;

        // 3. Create a template document with an HTML tag that will render the styled table.
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from JSON data:");
        // Insert the HTML expression tag.
        builder.Writeln("<<[model.HtmlTable] -html>>");
        templateDoc.Save(templatePath);

        // 4. Load the template and build the report using LINQ Reporting Engine.
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(loadedTemplate, model, "model");

        // 5. Save the final report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        loadedTemplate.Save(reportPath);
    }
}
