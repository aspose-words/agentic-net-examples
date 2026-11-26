using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LineItem
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = "";
    public List<LineItem> LineItems { get; set; } = new();
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling (required by Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel();
        model.Orders.Add(new Order
        {
            OrderId = 1001,
            CustomerName = "John Doe",
            LineItems = new List<LineItem>
            {
                new LineItem { ProductName = "Widget A", Quantity = 2, Price = 9.99m },
                new LineItem { ProductName = "Widget B", Quantity = 1, Price = 19.99m }
            }
        });
        model.Orders.Add(new Order
        {
            OrderId = 1002,
            CustomerName = "Jane Smith",
            LineItems = new List<LineItem>
            {
                new LineItem { ProductName = "Gadget X", Quantity = 5, Price = 4.99m }
            }
        });

        // Create the template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>   Customer: <<[order.CustomerName]>>");
        builder.Writeln("");
        builder.Writeln("<<foreach [item in order.LineItems]>>");
        builder.Writeln("  - <<[item.ProductName]>> x<<[item.Quantity]>> @ $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the model as the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
