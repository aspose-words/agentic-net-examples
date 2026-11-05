using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Sample data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result documents
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // Create sample data
        var order = new Order
        {
            Items = GenerateItems(150) // Adjust count to test threshold behavior
        };

        // Configure reflection optimization based on collection size threshold
        const int threshold = 100;
        ReportingEngine.UseReflectionOptimization = order.Items.Count > threshold;

        // Build the template document programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln($"Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var loadedTemplate = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, order, "order");

        // Save the generated report
        loadedTemplate.Save(resultPath);
    }

    // Helper method to generate a list of items
    private static List<Item> GenerateItems(int count)
    {
        var items = new List<Item>();
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Item { Index = i, Name = $"Product {i}" });
        }
        return items;
    }
}
