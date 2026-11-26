using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "Acme Corp",
            Items = GenerateItems(1500) // Adjust count to test threshold.
        };

        // Configure reflection optimization based on collection size.
        const int threshold = 1000;
        ReportingEngine.UseReflectionOptimization = order.Items.Count > threshold;

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static List<Item> GenerateItems(int count)
    {
        var items = new List<Item>(count);
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Item { Index = i, Name = $"Item {i}" });
        }
        return items;
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}
