using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LineItem
{
    public string Description { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}

public class Order
{
    public List<LineItem> Items { get; set; } = new();
    public string OrderNumber { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var order = new Order
        {
            OrderNumber = "ORD-001",
            Items = new List<LineItem>
            {
                new LineItem { Description = "Widget A", Price = 9.99m, Quantity = 3 },
                new LineItem { Description = "Widget B", Price = 14.50m, Quantity = 2 },
                new LineItem { Description = "Widget C", Price = 4.75m, Quantity = 5 }
            }
        };

        // Create template document
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln($"Order Number: <<[order.OrderNumber]>>");
        builder.Writeln();

        // Begin foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Create table header
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Total");
        builder.EndRow();

        // Table row for each item
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        // Arithmetic expression: price * quantity
        builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        // End table and foreach
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // Load template for report generation
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the order object as root named "order"
        engine.BuildReport(reportDoc, order, "order");

        // Ensure output directory exists
        var outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
