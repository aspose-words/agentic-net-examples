using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class Order
{
    public List<Item> Items { get; set; } = new();
    public string Customer { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            Customer = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 5 },
                new Item { Name = "Banana", Quantity = 3 },
                new Item { Name = "Orange", Quantity = 7 }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report for <<[order.Customer]>>");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Product: <<[Name]>> - Qty: <<[Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using contextual member access inside the data band.
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
