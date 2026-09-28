using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;          // Needed for Table type
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public decimal Discount { get; set; } // percentage
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new[]
        {
            new Item { Name = "Laptop", Price = 1200m, Discount = 10m },
            new Item { Name = "Smartphone", Price = 800m, Discount = 5m },
            new Item { Name = "Headphones", Price = 150m, Discount = 20m }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Deserialize JSON into the model.
        var model = new ReportModel
        {
            Items = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new()
        };

        // Create the LINQ Reporting template.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Discount Report");
        builder.Writeln();

        // Begin foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table inside the loop.
        Table table = builder.StartTable();

        // Table header (only once, before data rows).
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Discount (%)");
        builder.InsertCell();
        builder.Writeln("Discounted Price");
        builder.EndRow();

        // Data row template.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Discount]>>");
        builder.InsertCell();
        // Inline calculation: Price * (1 - Discount/100)
        builder.Writeln("<<[item.Price * (1 - item.Discount / 100)]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
