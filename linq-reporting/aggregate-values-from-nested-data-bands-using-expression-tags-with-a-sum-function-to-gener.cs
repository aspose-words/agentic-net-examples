using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Sample data
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Name = "Apple", Quantity = 3, Price = 0.5m },
                        new OrderItem { Name = "Banana", Quantity = 2, Price = 0.3m }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Name = "Orange", Quantity = 5, Price = 0.4m },
                        new OrderItem { Name = "Grapes", Quantity = 1, Price = 2.0m }
                    }
                }
            }
        };

        // Create template document
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin orders loop
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln();

        // Items – each item gets its own small table (header + one data row)
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Start table for a single item
        var table = builder.StartTable();

        // Header row
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Line Total");
        builder.EndRow();

        // Data row
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity * item.Price]>>");
        builder.EndRow();

        // End table for this item
        builder.EndTable();

        builder.Writeln("<</foreach>>"); // end items loop

        // Order total using sum expression
        builder.Writeln("Order Total: <<[order.Items.Sum(i => i.Quantity * i.Price)]>>");
        builder.Writeln("<</foreach>>"); // end orders loop

        // Grand total across all orders
        builder.Writeln("Grand Total: <<[Orders.Sum(o => o.Items.Sum(i => i.Quantity * i.Price))]>>");

        // Save template
        doc.Save(templatePath);

        // Load template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build report
        engine.BuildReport(reportDoc, model, "model");

        // Save final report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data models
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public List<OrderItem> Items { get; set; } = new();
}

public class OrderItem
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
