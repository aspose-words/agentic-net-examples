using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    CustomerName = "Alice",
                    Items = new List<Item>
                    {
                        new Item { Name = "Pen", Quantity = 2, UnitPrice = 1.5m },
                        new Item { Name = "Notebook", Quantity = 1, UnitPrice = 3.0m }
                    }
                },
                new Order
                {
                    CustomerName = "Bob",
                    Items = new List<Item>
                    {
                        new Item { Name = "Pencil", Quantity = 5, UnitPrice = 0.8m },
                        new Item { Name = "Eraser", Quantity = 2, UnitPrice = 0.5m }
                    }
                }
            }
        };

        // Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Begin outer foreach over Orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Table header for order items.
        builder.StartTable();
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Unit Price");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Inner foreach over order.Items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Total]>>");
        builder.EndRow();
        builder.Writeln("<</foreach>>");

        // Order total row – use the computed Order.Total property.
        builder.InsertCell(); builder.Writeln("Order Total:");
        builder.InsertCell(); builder.Writeln();
        builder.InsertCell(); builder.Writeln();
        builder.InsertCell(); builder.Writeln("<<[order.Total]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Grand total summary table – use the computed GrandTotal property.
        builder.Writeln("Grand Total Summary:");
        builder.StartTable();
        builder.InsertCell(); builder.Writeln("Grand Total:");
        builder.InsertCell(); builder.Writeln("<<[model.GrandTotal]>>");
        builder.EndRow();
        builder.EndTable();

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root model.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Grand total across all orders.
    public decimal GrandTotal
    {
        get
        {
            decimal sum = 0;
            foreach (var order in Orders)
                sum += order.Total;
            return sum;
        }
    }
}

// Order model with a computed Total property.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();

    // Sum of item totals for this order.
    public decimal Total
    {
        get
        {
            decimal sum = 0;
            foreach (var item in Items)
                sum += item.Total;
            return sum;
        }
    }
}

// Item model with a computed Total property.
public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }

    public decimal Total => Quantity * UnitPrice;
}
