using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new Order
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 5 },
                new Item { Name = "Banana", Quantity = 3 }
            }
        };

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);

        // Output result status.
        Console.WriteLine($"Report generation successful: {success}");
    }
}

// Data model classes.
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
