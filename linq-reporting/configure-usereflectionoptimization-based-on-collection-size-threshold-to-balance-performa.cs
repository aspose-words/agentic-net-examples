using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for encoding support.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>()
        };

        // Populate the collection with sample items.
        for (int i = 1; i <= 7; i++)
        {
            order.Items.Add(new Item { Name = $"Product {i}", Price = i * 10 });
        }

        // Define a threshold to decide whether to enable reflection optimization.
        const int threshold = 5;
        ReportingEngine.UseReflectionOptimization = order.Items.Count > threshold;

        // Create the template document.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
