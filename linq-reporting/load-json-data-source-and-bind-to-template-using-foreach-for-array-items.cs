using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        var sampleData = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20m },
                new Item { Name = "Banana", Price = 0.80m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into a strongly‑typed object.
        string jsonContent = File.ReadAllText(jsonPath);
        Order order = JsonConvert.DeserializeObject<Order>(jsonContent)!;

        // Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report for <<[order.CustomerName]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
