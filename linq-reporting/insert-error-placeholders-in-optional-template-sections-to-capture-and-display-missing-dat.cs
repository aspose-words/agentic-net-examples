using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "Acme Corp",
            Items = new()
            {
                new Item { Name = "Widget", Description = "Standard widget" },
                new Item { Name = "Gadget", Description = null }, // Missing description to trigger <<error>>
                new Item { Name = "Doohickey", Description = "Multi‑purpose tool" }
            }
        };

        // Create template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<<if [item.Description]>> Description: <<[item.Description]>> <</if>>");
        // Insert <<error>> placeholder to capture missing data warnings.
        builder.Writeln("<<error>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Configure reporting engine to show inline error messages.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
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
    public string? Description { get; set; }
}
