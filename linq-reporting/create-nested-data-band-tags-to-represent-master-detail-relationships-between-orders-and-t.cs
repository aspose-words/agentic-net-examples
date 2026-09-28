using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = "";
    public List<LineItem> LineItems { get; set; } = new();
}

public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Orders = new()
            {
                new Order
                {
                    OrderId = 1,
                    CustomerName = "Alice",
                    LineItems = new()
                    {
                        new LineItem { ProductName = "Widget", Quantity = 2, Price = 9.99m },
                        new LineItem { ProductName = "Gadget", Quantity = 1, Price = 19.99m }
                    }
                },
                new Order
                {
                    OrderId = 2,
                    CustomerName = "Bob",
                    LineItems = new()
                    {
                        new LineItem { ProductName = "Thingamajig", Quantity = 5, Price = 4.99m }
                    }
                }
            }
        };

        // Create template with nested data band tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>   Customer: <<[order.CustomerName]>>");
        builder.Writeln("Line Items:");
        builder.Writeln("<<foreach [item in order.LineItems]>>");
        builder.Writeln("- <<[item.ProductName]>>   Qty: <<[item.Quantity]>>   Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
