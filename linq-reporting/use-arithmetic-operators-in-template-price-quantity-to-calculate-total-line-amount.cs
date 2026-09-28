using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            Customer = "John Doe",
            Items = new()
            {
                new Item { Name = "Apple", Price = 0.5m, Quantity = 4 },
                new Item { Name = "Banana", Price = 0.3m, Quantity = 6 },
                new Item { Name = "Cherry", Price = 1.2m, Quantity = 10 }
            }
        };

        // Create a template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Customer line.
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln();

        // Begin foreach loop for items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Start table inside the loop (the table will be generated once per order).
        builder.StartTable();

        // Table header.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Total");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        // Arithmetic expression: price * quantity.
        builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the result.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class Order
{
    public string Customer { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
