using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new ReportModel
        {
            Order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Price = 1.20 },
                    new Item { Name = "Banana", Price = 0.80 }
                },
                // EmptyTag is null to produce an empty paragraph after processing.
                EmptyTag = null
                // MissingProperty is intentionally omitted to trigger an inline error.
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure the reporting engine with both options.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages
        };

        // Build the report; the returned flag indicates success when InlineErrorMessages is set.
        bool success = engine.BuildReport(doc, model, "order");

        // Save the generated report.
        var outputPath = "ReportOutput.docx";
        doc.Save(outputPath);

        // Output simple status (no interactive prompts).
        Console.WriteLine($"Report built successfully: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header with a valid field.
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Loop over items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Tag that will be empty (EmptyTag is null).
        builder.Writeln("<<[order.EmptyTag]>>");

        // Tag referencing a missing property to demonstrate inline error messages.
        builder.Writeln("Missing: <<[order.MissingProperty]>>");

        doc.Save(filePath);
    }
}

// Root wrapper class to align with BuildReport(rootObject, "order").
public class ReportModel
{
    public Order Order { get; set; } = new();
}

// Sample order class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
    public string? EmptyTag { get; set; }
    // Note: MissingProperty is intentionally not defined.
}

// Sample item class.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
