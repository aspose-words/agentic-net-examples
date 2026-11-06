using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data file.
        const string jsonPath = "data.json";
        var sampleData = new List<Product>
        {
            new Product { Name = "Apple", Price = 0.99m },
            new Product { Name = "Banana", Price = 0.59m },
            new Product { Name = "Cherry", Price = 2.99m }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON data.
        var json = File.ReadAllText(jsonPath);
        var products = JsonConvert.DeserializeObject<List<Product>>(json) ?? new List<Product>();

        // Generate styled HTML table from the data.
        var htmlBuilder = new StringBuilder();
        htmlBuilder.AppendLine("<table style=\"border-collapse:collapse;width:100%;font-family:Arial;\">");
        htmlBuilder.AppendLine("<tr>");
        htmlBuilder.AppendLine("<th style=\"border:1px solid #dddddd;padding:8px;background:#f2f2f2;\">Product</th>");
        htmlBuilder.AppendLine("<th style=\"border:1px solid #dddddd;padding:8px;background:#f2f2f2;\">Price (USD)</th>");
        htmlBuilder.AppendLine("</tr>");
        foreach (var p in products)
        {
            htmlBuilder.AppendLine("<tr>");
            htmlBuilder.AppendLine($"<td style=\"border:1px solid #dddddd;padding:8px;\">{HtmlEncoder.Default.Encode(p.Name)}</td>");
            htmlBuilder.AppendLine($"<td style=\"border:1px solid #dddddd;padding:8px;text-align:right;\">{p.Price:F2}</td>");
            htmlBuilder.AppendLine("</tr>");
        }
        htmlBuilder.AppendLine("</table>");

        // Prepare the model for LINQ Reporting.
        var model = new ReportModel
        {
            HtmlTable = htmlBuilder.ToString()
        };

        // Create the template document programmatically.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from JSON data:");
        builder.Writeln("<<[model.HtmlTable] -html>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root model class used by the reporting engine.
public class ReportModel
{
    // HTML snippet that will be inserted into the document.
    public string HtmlTable { get; set; } = string.Empty;
}

// Simple data class representing a product.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
