using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document with LINQ tags.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Master band: orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln();

        // Detail band: line items.
        builder.Writeln("Line Items:");
        builder.Writeln("<<foreach [item in order.LineItems]>>");

        Table table = builder.StartTable();
        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>"); // end inner foreach (line items)
        builder.Writeln(); // blank line between orders
        builder.Writeln("<</foreach>>"); // end outer foreach (orders)

        // Save the template.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare sample data.
        // -------------------------------------------------
        ReportModel model = new()
        {
            Orders = new()
            {
                new Order
                {
                    OrderId = 1001,
                    CustomerName = "Alice Johnson",
                    OrderDate = DateTime.Today,
                    LineItems = new()
                    {
                        new LineItem { ProductName = "Widget A", Quantity = 3, Price = 19.99m },
                        new LineItem { ProductName = "Gadget B", Quantity = 1, Price = 99.50m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    CustomerName = "Bob Smith",
                    OrderDate = DateTime.Today.AddDays(-1),
                    LineItems = new()
                    {
                        new LineItem { ProductName = "Widget C", Quantity = 5, Price = 9.75m },
                        new LineItem { ProductName = "Gadget D", Quantity = 2, Price = 49.99m },
                        new LineItem { ProductName = "Accessory E", Quantity = 10, Price = 2.99m }
                    }
                }
            }
        };

        // -------------------------------------------------
        // 3. Build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// -------------------------------------------------
// Data model classes.
// -------------------------------------------------
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public List<LineItem> LineItems { get; set; } = new();
}

public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
