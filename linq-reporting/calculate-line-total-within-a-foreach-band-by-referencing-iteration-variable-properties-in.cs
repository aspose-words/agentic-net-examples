using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln();

        // Header row for the table.
        builder.Writeln("<<foreach [line in OrderLines]>>"); // Open foreach band.

        // Start the table that will contain all line items.
        Table table = builder.StartTable();

        // Table header.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Unit Price");
        builder.InsertCell();
        builder.Writeln("Line Total");
        builder.EndRow();

        // Row for each line item.
        builder.InsertCell();
        builder.Writeln("<<[line.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[line.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[line.UnitPrice]>>");
        builder.InsertCell();
        // Calculate line total directly in the expression.
        builder.Writeln("<<[line.Quantity * line.UnitPrice]>>");
        builder.EndRow();

        // End the table and the foreach band.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Prepare sample data.
        Order order = new Order
        {
            OrderLines = new List<LineItem>
            {
                new LineItem { ProductName = "Apple", Quantity = 3, UnitPrice = 0.5m },
                new LineItem { ProductName = "Banana", Quantity = 5, UnitPrice = 0.3m },
                new LineItem { ProductName = "Cherry", Quantity = 2, UnitPrice = 1.2m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, order, "order");

        // Save the generated report.
        const string outputPath = "output.docx";
        report.Save(outputPath);
    }
}

// Root data model.
public class Order
{
    public List<LineItem> OrderLines { get; set; } = new();
}

// Item model used inside the foreach band.
public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
