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
        Order order = new()
        {
            CustomerName = "John Doe",
            Items = new()
            {
                new Item { Index = 1, Name = "Widget A" },
                new Item { Index = 2, Name = "Widget B" },
                new Item { Index = 3, Name = "Widget C" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Disable reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
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
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
