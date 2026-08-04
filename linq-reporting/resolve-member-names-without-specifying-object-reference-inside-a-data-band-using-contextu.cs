using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

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
                new Item { Name = "Apple", Price = 1.20m },
                new Item { Name = "Banana", Price = 0.80m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer: <<[CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - <<[Name]>> : $<<[Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
