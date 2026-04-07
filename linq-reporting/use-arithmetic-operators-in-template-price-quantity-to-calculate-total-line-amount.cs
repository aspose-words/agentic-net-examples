using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data.
        var order = new Order
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Price = 0.50m, Quantity = 4 },
                new Item { Name = "Banana", Price = 0.30m, Quantity = 6 },
                new Item { Name = "Orange", Price = 0.80m, Quantity = 3 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Invoice");
        builder.Writeln();

        // Begin the foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table creation inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Data row – each cell contains a tag that will be evaluated per item.
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        // Arithmetic expression: price * quantity.
        builder.InsertCell(); builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the order object as the data source.
        engine.BuildReport(reportDoc, order);

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (must be public with public properties).
// ---------------------------------------------------------------------
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
