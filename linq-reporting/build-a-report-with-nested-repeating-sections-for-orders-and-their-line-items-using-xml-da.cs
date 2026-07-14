using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a Word template with LINQ Reporting tags.
        var templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders Report");
        builder.Writeln("");

        // Outer foreach for orders.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Items:");
        // Inner foreach for line items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.ProductName]>>: <<[item.Quantity]>> x $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template.
        var reportDoc = new Document(templatePath);

        // Prepare data model.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new()
                {
                    CustomerName = "John Doe",
                    OrderDate = new DateTime(2023, 1, 15),
                    Items = new List<Item>
                    {
                        new() { ProductName = "Widget A", Quantity = 2, Price = 9.99m },
                        new() { ProductName = "Widget B", Quantity = 1, Price = 19.99m }
                    }
                },
                new()
                {
                    CustomerName = "Jane Smith",
                    OrderDate = new DateTime(2023, 2, 5),
                    Items = new List<Item>
                    {
                        new() { ProductName = "Gadget X", Quantity = 5, Price = 4.50m }
                    }
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
