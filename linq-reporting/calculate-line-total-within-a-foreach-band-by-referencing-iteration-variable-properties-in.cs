using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new List<LineItem>
            {
                new LineItem { Description = "Widget A", Quantity = 3, UnitPrice = 19.99m },
                new LineItem { Description = "Widget B", Quantity = 5, UnitPrice = 9.50m },
                new LineItem { Description = "Service C", Quantity = 2, UnitPrice = 150.00m }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with customer name.
        builder.Writeln("Invoice for: <<[order.CustomerName]>>");
        builder.Writeln();
        // Table header.
        builder.Writeln("Item\tQty\tPrice\tTotal");
        // Start a foreach band over the collection of line items.
        builder.Writeln("<<foreach [line in Items]>>");
        // Each row displays description, quantity, unit price and calculated line total.
        builder.Writeln("<<[line.Description]>>\t<<[line.Quantity]>>\t<<[line.UnitPrice]>>\t<<[line.Quantity * line.UnitPrice]>>");
        // End of the foreach band.
        builder.Writeln("<</foreach>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object is 'order', so we reference it in the template with the name "order".
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Order
{
    // Initialize to avoid nullable warnings.
    public string CustomerName { get; set; } = string.Empty;
    public List<LineItem> Items { get; set; } = new();
}

public class LineItem
{
    public string Description { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
