using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Url = "https://example.com", LinkText = "Example Site" },
                new Item { Url = "https://test.com", LinkText = null } // No display text.
            }
        };

        // -----------------------------------------------------------------
        // Step 1: Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write a heading.
        builder.Writeln("Link Tag Demonstration");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Link with explicit display text (if provided).
        builder.Writeln("Explicit: <<link [item.Url] [item.LinkText]>>");
        builder.Writeln();

        // Link without a second expression – the URL itself becomes the display text.
        builder.Writeln("Implicit: <<link [item.Url]>>");
        builder.Writeln();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "LinkTemplate.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "LinkReport.docx";
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Url { get; set; } = string.Empty;
    public string? LinkText { get; set; }
}
