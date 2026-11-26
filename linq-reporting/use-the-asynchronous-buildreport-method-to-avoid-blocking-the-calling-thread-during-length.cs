using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    // Entry point – asynchronous to allow awaiting the report generation.
    static async Task Main(string[] args)
    {
        // Register code page provider for any legacy encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple header with a data field.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Loop over the collection of items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data model.
        // -----------------------------------------------------------------
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Cherry", Quantity = 12 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report asynchronously.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Wrap the synchronous BuildReport call in Task.Run to avoid blocking.
        bool success = await Task.Run(() => engine.BuildReport(reportDoc, order, "order"));

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);

        // Indicate the result.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// ---------------------------------------------------------------------
// Data model classes – must be public with public properties.
// ---------------------------------------------------------------------
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
