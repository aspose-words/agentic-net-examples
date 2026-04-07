using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Order Report");
        builder.Writeln();

        // Iterate over the orders collection.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        // Calculate order total on the fly using LINQ inside the tag.
        builder.Writeln("Order Total: <<[order.Items.Sum(i => i.Quantity * i.UnitPrice)]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Summary section – aggregate totals of all orders.
        builder.Writeln("Summary Total: <<[model.Orders.Sum(o => o.Items.Sum(i => i.Quantity * i.UnitPrice))]>>");

        // Prepare sample data.
        ReportModel model = new()
        {
            Orders = new List<Order>
            {
                new()
                {
                    CustomerName = "Alice",
                    Items = new List<Item>
                    {
                        new() { Name = "Pen", Quantity = 10, UnitPrice = 1.20m },
                        new() { Name = "Notebook", Quantity = 5, UnitPrice = 3.50m }
                    }
                },
                new()
                {
                    CustomerName = "Bob",
                    Items = new List<Item>
                    {
                        new() { Name = "Pencil", Quantity = 20, UnitPrice = 0.80m },
                        new() { Name = "Eraser", Quantity = 2, UnitPrice = 1.00m }
                    }
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ReportOutput.docx");
    }
}

// Root data model containing a collection of orders.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Represents a single order.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Represents an item within an order.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
