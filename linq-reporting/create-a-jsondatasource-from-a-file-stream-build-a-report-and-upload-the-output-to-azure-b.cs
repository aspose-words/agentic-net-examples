using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main(string[] args)
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Cherry", Quantity = 2 }
            }
        };

        // Serialize model to JSON file.
        const string jsonFilePath = "data.json";
        File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(new { order }, Formatting.Indented));

        // Create a Word template with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template.
        var templateDoc = new Document(templatePath);

        // Load JSON data from a file stream and build the report.
        using (FileStream jsonStream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
        {
            var jsonDataSource = new JsonDataSource(jsonStream);
            var engine = new ReportingEngine();

            // Build the report using the root name "order".
            engine.BuildReport(templateDoc, jsonDataSource, "order");
        }

        // Save the generated report.
        const string reportPath = "report.docx";
        templateDoc.Save(reportPath);

        // NOTE: Azure Blob Storage upload requires the Azure.Storage.Blobs package.
        // The example omits the upload step to keep the project self‑contained and compilable.
        Console.WriteLine($"Report generated and saved to '{reportPath}'.");
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
