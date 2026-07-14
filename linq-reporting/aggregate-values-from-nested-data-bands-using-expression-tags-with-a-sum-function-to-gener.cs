using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class Order
{
    public int OrderId { get; set; }
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new()
                {
                    OrderId = 1,
                    Items = new List<Item>
                    {
                        new() { Product = "Apple", Quantity = 3, Price = 0.5m },
                        new() { Product = "Banana", Quantity = 2, Price = 0.3m }
                    }
                },
                new()
                {
                    OrderId = 2,
                    Items = new List<Item>
                    {
                        new() { Product = "Apple", Quantity = 1, Price = 0.5m },
                        new() { Product = "Orange", Quantity = 5, Price = 0.4m }
                    }
                }
            }
        };

        // Create template document
        var templatePath = "ReportTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order #: <<[order.OrderId]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Product]>>: Qty <<[item.Quantity]>>, Price <<[item.Price]>>, Total <<[item.Price * item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Order Total: <<[order.Items.Sum(i => i.Price * i.Quantity)]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Overall Summary:");
        builder.StartTable();
        builder.InsertCell(); builder.Writeln("Total Amount");
        builder.InsertCell(); builder.Writeln("<<[Orders.SelectMany(o => o.Items).Sum(i => i.Price * i.Quantity)]>>");
        builder.EndRow();
        builder.EndTable();

        doc.Save(templatePath);

        // Load template and build report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report
        var outputPath = "ReportOutput.docx";
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
