using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with unique bookmark names
        var model = new ReportModel
        {
            Items = new()
        };

        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new Row
            {
                Name = $"Item {i}",
                Description = $"Description for item {i}",
                BookmarkName = $"Row_{i}"
            });
        }

        // Create a template document programmatically
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Report with Unique Bookmarks");
        builder.Writeln();

        // Begin foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table inside the foreach block
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name (Bookmark)");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.EndRow();

        // Data row (template)
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>> <<[item.Name]>> <</bookmark>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.EndRow();

        // End table
        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the generated document
        string outputPath = Path.Combine(outputDir, "ReportWithBookmarks.docx");
        doc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Root data model
public class ReportModel
{
    public List<Row> Items { get; set; } = new();
}

// Row data model
public class Row
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
