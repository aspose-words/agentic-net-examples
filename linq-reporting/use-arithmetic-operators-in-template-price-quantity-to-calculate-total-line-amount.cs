using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data
        Order order = new()
        {
            OrderNumber = "ORD-001",
            Items = new()
            {
                new Item { Name = "Apple",  Price = 0.5m, Quantity = 10 },
                new Item { Name = "Banana", Price = 0.3m, Quantity = 15 },
                new Item { Name = "Cherry", Price = 1.2m, Quantity = 5 }
            }
        };

        // Create template document
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Report title
        builder.Writeln($"Order Report: <<[order.OrderNumber]>>");
        builder.Writeln();

        // Begin foreach over Items (must be placed before the table)
        builder.Writeln("<<foreach [item in Items]>>");

        // Table definition
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell(); builder.Writeln("Product");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Data row (will be repeated for each item)
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        // Close table
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load template for report generation
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes
public class Order
{
    public string OrderNumber { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
