using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ReportModel
{
    public string HtmlTable { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- 1. Prepare sample JSON data ----------
        string jsonPath = "data.json";
        var sampleItems = new List<Item>
        {
            new Item { Name = "Apple", Price = 1.20m },
            new Item { Name = "Banana", Price = 0.80m },
            new Item { Name = "Cherry", Price = 2.50m }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleItems, Formatting.Indented));

        // ---------- 2. Load JSON and build HTML table ----------
        var items = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new List<Item>();
        var sb = new StringBuilder();
        sb.AppendLine("<table style='border:1px solid black;border-collapse:collapse;width:50%;'>");
        sb.AppendLine("<tr><th style='border:1px solid black;background:#D3D3D3;'>Name</th><th style='border:1px solid black;background:#D3D3D3;'>Price</th></tr>");
        foreach (var it in items)
        {
            sb.AppendLine($"<tr><td style='border:1px solid black;padding:5px;'>{System.Net.WebUtility.HtmlEncode(it.Name)}</td><td style='border:1px solid black;padding:5px;'>{it.Price:C}</td></tr>");
        }
        sb.AppendLine("</table>");

        var model = new ReportModel { HtmlTable = sb.ToString() };

        // ---------- 3. Create template document with HTML switch ----------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[model.HtmlTable] -html>>");
        templateDoc.Save(templatePath);

        // ---------- 4. Load template and build report ----------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // ---------- 5. Save final document ----------
        string resultPath = "ReportResult.docx";
        reportDoc.Save(resultPath);
    }
}
