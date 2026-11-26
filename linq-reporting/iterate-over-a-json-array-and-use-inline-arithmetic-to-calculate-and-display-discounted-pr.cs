using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public decimal Discount { get; set; } // e.g., 0.15 for 15%
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"
        [
            { ""Name"": ""Laptop"", ""Price"": 1200.00, ""Discount"": 0.10 },
            { ""Name"": ""Smartphone"", ""Price"": 800.00, ""Discount"": 0.15 },
            { ""Name"": ""Headphones"", ""Price"": 150.00, ""Discount"": 0.20 }
        ]";

        // Deserialize JSON into model.
        var products = JsonConvert.DeserializeObject<List<Product>>(json) ?? new List<Product>();
        var model = new ReportModel { Products = products };

        // Create template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Header table (static, not repeated).
        Table headerTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Discount (%)");
        builder.InsertCell(); builder.Writeln("Discounted Price");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach loop over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Data rows table (repeated for each product).
        Table dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[p.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Discount * 100]>>%");
        builder.InsertCell(); builder.Writeln("<<[p.Price - (p.Price * p.Discount)]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as root named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
