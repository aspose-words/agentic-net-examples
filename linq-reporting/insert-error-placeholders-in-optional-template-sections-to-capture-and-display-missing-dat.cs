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
        // Register code page provider for possible encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple header.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over order items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>");
        // Price may be missing; <<error>> will display a warning when InlineErrorMessages is enabled.
        builder.Writeln("Price: <<[item.Price]>> <<error>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare sample data with a missing Price for one item.
        // -----------------------------------------------------------------
        Order order = new()
        {
            CustomerName = "John Doe",
            Items = new()
            {
                new Item { Name = "Apple", Quantity = 5, Price = 1.20m },
                new Item { Name = "Banana", Quantity = 3, Price = null } // Missing price triggers <<error>>
            }
        };

        // -----------------------------------------------------------------
        // Build the report with inline error messages enabled.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(outputPath);

        // Output the success flag (true if no fatal errors occurred).
        Console.WriteLine($"Report generation success: {success}");
    }
}

// ---------------------------------------------------------------------
// Public data model classes.
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
    public decimal? Price { get; set; }
}
