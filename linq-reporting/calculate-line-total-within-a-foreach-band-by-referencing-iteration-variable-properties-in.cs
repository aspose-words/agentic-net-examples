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
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Order Details");
        builder.Writeln();

        // Begin the foreach band – iterate over order.Items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Table with header and data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Unit Price");
        builder.InsertCell(); builder.Writeln("Line Total");
        builder.EndRow();

        // Data row – will be repeated for each item.
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
        // Calculate line total directly in the expression tag.
        builder.InsertCell(); builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
        builder.EndRow();

        // Close the table.
        builder.EndTable();

        // End the foreach band.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample order with a few items.
        Order order = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Quantity = 3, UnitPrice = 0.75m },
                new Item { Name = "Banana", Quantity = 5, UnitPrice = 0.50m },
                new Item { Name = "Cherry", Quantity = 2, UnitPrice = 2.00m }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template is "order".
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
