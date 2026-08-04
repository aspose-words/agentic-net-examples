using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a new blank document and a builder to insert LINQ Reporting tags
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Master‑detail tags: outer foreach for orders, inner foreach for line items
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order #: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.ProductName]>> (Qty: <<[item.Quantity]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data
        ReportModel model = new ReportModel
        {
            Orders = GetSampleOrders()
        };

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document
        doc.Save("NestedReport.docx");
    }

    private static List<Order> GetSampleOrders()
    {
        return new List<Order>
        {
            new Order
            {
                OrderId = 1001,
                CustomerName = "Alice",
                Items = new List<LineItem>
                {
                    new LineItem { ProductName = "Apple", Quantity = 3 },
                    new LineItem { ProductName = "Banana", Quantity = 5 }
                }
            },
            new Order
            {
                OrderId = 1002,
                CustomerName = "Bob",
                Items = new List<LineItem>
                {
                    new LineItem { ProductName = "Orange", Quantity = 2 },
                    new LineItem { ProductName = "Grapes", Quantity = 1 },
                    new LineItem { ProductName = "Mango", Quantity = 4 }
                }
            }
        };
    }
}

// Wrapper class used as the root data source for the report
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Master object representing an order
public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<LineItem> Items { get; set; } = new();
}

// Detail object representing a line item in an order
public class LineItem
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
