using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln();

        // Begin foreach band that will repeat the table rows.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create the table inside the foreach band.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Qty");
        builder.InsertCell();
        builder.Writeln("Unit Price");
        builder.InsertCell();
        builder.Writeln("Line Total");
        builder.EndRow();

        // Data row – will be repeated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.UnitPrice]>>");
        builder.InsertCell();
        // Calculate line total using an expression tag.
        builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
        builder.EndRow();

        // End the table and the foreach band.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Sample data.
        var order = new Order
        {
            Items = new List<LineItem>
            {
                new LineItem { ProductName = "Apple",  Quantity = 3, UnitPrice = 0.5m },
                new LineItem { ProductName = "Banana", Quantity = 5, UnitPrice = 0.3m },
                new LineItem { ProductName = "Cherry", Quantity = 2, UnitPrice = 1.2m }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Order
{
    public List<LineItem> Items { get; set; } = new();
}

public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
