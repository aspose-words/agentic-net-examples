using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for encoding support required by Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        ReportModel model = new()
        {
            Title = "Sample LINQ Reporting",
            HtmlSnippet = "<b>Bold Text</b>",
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Title placeholder.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("<<[model.Title]>>");

        // Insert HTML snippet.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Begin foreach loop.
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Table definition inside the loop.
        Table table = builder.StartTable();

        // Header row (only once, so we add it before the first data row).
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row template.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End of table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(template, model, "model");

        // Serialize the generated report to a memory stream.
        using MemoryStream reportStream = new();
        template.Save(reportStream, SaveFormat.Docx);
        reportStream.Position = 0; // Reset for further processing.

        // Demonstrate that the report was generated.
        Console.WriteLine($"Report generated successfully: {success}");
        Console.WriteLine($"Report size (bytes): {reportStream.Length}");
    }

    // Data model classes.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public string HtmlSnippet { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }
}
