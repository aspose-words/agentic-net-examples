using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for Table type

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Build the template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Order Report");
        builder.Writeln();

        // Outer foreach over orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln();

        // Header for items table.
        builder.Writeln("Items:");
        // Inner foreach over items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Start table for each order's items.
        Table itemsTable = builder.StartTable();

        // Table header.
        builder.InsertCell(); builder.Writeln("Product");
        builder.InsertCell(); builder.Writeln("Qty");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Total]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Close inner foreach.
        builder.Writeln("<</foreach>>");

        // Order total using a sum expression.
        builder.Writeln("Order Total: <<[order.Items.Sum(i => i.Total)]>>");
        builder.Writeln();

        // Close outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Create sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    Id = 1001,
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple", Quantity = 3, Price = 0.5m },
                        new Item { Name = "Banana", Quantity = 5, Price = 0.3m }
                    }
                },
                new Order
                {
                    Id = 1002,
                    Items = new List<Item>
                    {
                        new Item { Name = "Orange", Quantity = 2, Price = 0.8m },
                        new Item { Name = "Grapes", Quantity = 1, Price = 2.5m },
                        new Item { Name = "Mango", Quantity = 4, Price = 1.2m }
                    }
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}

// Wrapper root object for the template.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Order containing a collection of items.
public class Order
{
    public int Id { get; set; }
    public List<Item> Items { get; set; } = new();
}

// Individual line item.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Price { get; set; }

    // Calculated total for the line item.
    public decimal Total => Quantity * Price;
}
