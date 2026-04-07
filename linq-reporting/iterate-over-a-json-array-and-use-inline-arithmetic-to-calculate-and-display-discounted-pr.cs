using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Simple data model for an item.
    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public decimal Discount { get; set; } // e.g., 0.15 for 15%
    }

    public static void Main()
    {
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var items = new List<Item>
        {
            new Item { Name = "Apple",  Price = 1.20m, Discount = 0.10m },
            new Item { Name = "Banana", Price = 0.80m, Discount = 0.05m },
            new Item { Name = "Cherry", Price = 2.50m, Discount = 0.20m }
        };

        // Serialize data to JSON and write to a temporary file.
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "data.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(items, Formatting.Indented));

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        CreateTemplate(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Load JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        // The top‑level JSON element is an array, so we reference it with the name we provide here.
        engine.BuildReport(doc, jsonDataSource, "items");

        // Save the generated report.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "report.docx");
        doc.Save(reportPath);
    }

    // Creates a simple Word document containing the required LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header.
        builder.Writeln("Discounted Price Report");
        builder.Writeln();

        // Begin foreach over the JSON array named 'items'.
        builder.Writeln("<<foreach [item in items]>>");

        // Output each item's details and calculate discounted price inline.
        builder.Writeln("Product: <<[item.Name]>>");
        builder.Writeln("Original Price: $<<[item.Price]>>");
        builder.Writeln("Discount: <<[item.Discount * 100]>>%"); // show discount as percentage
        builder.Writeln("Discounted Price: $<<[item.Price * (1 - item.Discount)]>>");
        builder.Writeln(); // blank line between items

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}
