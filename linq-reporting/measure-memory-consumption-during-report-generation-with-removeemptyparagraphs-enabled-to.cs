using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item 1", Description = "First item description." },
                new Item { Name = "Item 2", Description = "" },               // Empty description will lead to an empty paragraph.
                new Item { Name = "Item 3", Description = "Third item description." }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report generated with Aspose.Words LINQ Reporting");
        builder.Writeln(); // Blank line.

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Measure memory before building the report.
        long memoryBefore = GC.GetTotalMemory(true);

        // Build the report. The root object name is "model" as referenced in the template.
        engine.BuildReport(doc, model, "model");

        // Measure memory after building the report.
        long memoryAfter = GC.GetTotalMemory(true);

        // Output memory consumption details.
        Console.WriteLine($"Memory before report generation: {memoryBefore:N0} bytes");
        Console.WriteLine($"Memory after  report generation: {memoryAfter:N0} bytes");
        Console.WriteLine($"Memory difference: {memoryAfter - memoryBefore:N0} bytes");

        // Save the generated report.
        doc.Save("ReportOutput.docx");
    }
}

// Wrapper class that matches the root object name used in BuildReport.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple data item class.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
