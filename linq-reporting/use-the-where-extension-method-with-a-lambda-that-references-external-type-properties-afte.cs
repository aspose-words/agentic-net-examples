using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public double Price { get; set; }
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

// External type whose static members will be used inside the LINQ expression.
public static class ExternalHelper
{
    public static double MinPrice => 20.0;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Pen", Price = 5.0 },
                new Item { Name = "Notebook", Price = 12.5 },
                new Item { Name = "Backpack", Price = 45.0 },
                new Item { Name = "Calculator", Price = 30.0 }
            }
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items with price > <<[ExternalHelper.MinPrice]>>:");
        builder.Writeln("<<foreach [item in order.Items.Where(i => i.Price > ExternalHelper.MinPrice)]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(ExternalHelper)); // Register external type.

        // Build the report using the root object name "order".
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
