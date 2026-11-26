using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    CustomerName = "Alice Johnson",
                    Items = new List<LineItem>
                    {
                        new LineItem { ProductName = "Laptop", Quantity = 1 },
                        new LineItem { ProductName = "Mouse", Quantity = 2 }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    CustomerName = "Bob Smith",
                    Items = new List<LineItem>
                    {
                        new LineItem { ProductName = "Desk Chair", Quantity = 1 },
                        new LineItem { ProductName = "Monitor", Quantity = 2 },
                        new LineItem { ProductName = "Keyboard", Quantity = 1 }
                    }
                }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.ProductName]>> (Qty: <<[item.Quantity]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the result.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);

        // Simple verification output.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to '{outputPath}'.");
    }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<LineItem> Items { get; set; } = new();
}

public class LineItem
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
