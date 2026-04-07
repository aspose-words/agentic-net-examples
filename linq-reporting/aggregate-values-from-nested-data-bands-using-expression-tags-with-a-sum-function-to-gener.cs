using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer loop – iterate over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln();

        // Inner loop – iterate over items of the current order.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Sum of the current order's items (computed in the data model).
        builder.Writeln("Order Total: $<<[order.OrderTotal]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Overall summary – sum of all orders (computed in the data model).
        builder.Writeln("Overall Total: $<<[model.OverallTotal]>>");

        // Build the data model.
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    CustomerName = "Alice",
                    Items = new List<Item>
                    {
                        new Item { Name = "Pen", Price = 1.20m },
                        new Item { Name = "Notebook", Price = 3.45m }
                    }
                },
                new Order
                {
                    CustomerName = "Bob",
                    Items = new List<Item>
                    {
                        new Item { Name = "Pencil", Price = 0.80m },
                        new Item { Name = "Eraser", Price = 0.50m },
                        new Item { Name = "Ruler", Price = 1.10m }
                    }
                }
            }
        };

        // Run the LINQ reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("OrdersReport.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Overall total calculated from all orders.
    public decimal OverallTotal => Orders.Sum(o => o.OrderTotal);
}

// Order with a collection of items.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();

    // Total price for this order.
    public decimal OrderTotal => Items.Sum(i => i.Price);
}

// Individual line item.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
