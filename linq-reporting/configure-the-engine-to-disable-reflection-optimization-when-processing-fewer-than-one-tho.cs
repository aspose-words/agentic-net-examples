using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data with fewer than 1000 records.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new()
        };
        for (int i = 1; i <= 5; i++)
        {
            order.Items.Add(new Item { Name = $"Item {i}", Price = i * 10m });
        }

        // Disable reflection optimization when processing fewer than 1000 records.
        ReportingEngine.UseReflectionOptimization = order.Items.Count >= 1000;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
