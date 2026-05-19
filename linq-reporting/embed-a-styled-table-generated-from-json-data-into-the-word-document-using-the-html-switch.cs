using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Net;

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();

    // Generates a styled HTML table from the Products collection.
    public string HtmlTable => GenerateHtmlTable();

    private string GenerateHtmlTable()
    {
        var sb = new StringBuilder();

        sb.Append("<table style='border-collapse:collapse;width:100%;font-family:Arial;'>");
        sb.Append("<tr>");
        sb.Append("<th style='border:1px solid #000;background:#D3D3D3;padding:5px;'>Name</th>");
        sb.Append("<th style='border:1px solid #000;background:#D3D3D3;padding:5px;'>Price</th>");
        sb.Append("<th style='border:1px solid #000;background:#D3D3D3;padding:5px;'>Quantity</th>");
        sb.Append("</tr>");

        foreach (var p in Products)
        {
            sb.Append("<tr>");
            sb.Append($"<td style='border:1px solid #000;padding:5px;'>{WebUtility.HtmlEncode(p.Name)}</td>");
            sb.Append($"<td style='border:1px solid #000;padding:5px;'>{p.Price:C}</td>");
            sb.Append($"<td style='border:1px solid #000;padding:5px;'>{p.Quantity}</td>");
            sb.Append("</tr>");
        }

        sb.Append("</table>");
        return sb.ToString();
    }
}

public class Program
{
    private const string TemplatePath = "Template.docx";
    private const string JsonPath = "data.json";
    private const string OutputPath = "Report.docx";

    public static void Main()
    {
        // 1. Create sample JSON data.
        CreateSampleJson();

        // 2. Load JSON data into the model.
        var model = new ReportModel
        {
            Products = JsonConvert.DeserializeObject<List<Product>>(File.ReadAllText(JsonPath)) ?? new List<Product>()
        };

        // 3. Build the template document with an HTML placeholder.
        CreateTemplate();

        // 4. Load the template and run the reporting engine.
        var doc = new Document(TemplatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the final report.
        doc.Save(OutputPath);
    }

    private static void CreateSampleJson()
    {
        var sample = new List<Product>
        {
            new Product { Name = "Apple", Price = 0.99m, Quantity = 10 },
            new Product { Name = "Banana", Price = 0.59m, Quantity = 20 },
            new Product { Name = "Cherry", Price = 2.99m, Quantity = 15 }
        };

        var json = JsonConvert.SerializeObject(sample, Formatting.Indented);
        File.WriteAllText(JsonPath, json, Encoding.UTF8);
    }

    private static void CreateTemplate()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report generated from JSON data:");
        // Insert the HTML expression tag that will be replaced by the generated table.
        builder.Writeln("<<[model.HtmlTable] -html>>");

        doc.Save(TemplatePath);
    }
}
