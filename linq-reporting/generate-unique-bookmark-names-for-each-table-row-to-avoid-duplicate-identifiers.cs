using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Prepare sample data with unique bookmark names.
        var model = new ReportModel
        {
            Items = new List<RowItem>()
        };

        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new RowItem
            {
                Index = i,
                Name = $"Item {i}",
                // Ensure each bookmark name is unique.
                BookmarkName = $"Row_{i}"
            });
        }

        // Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin the foreach block.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build a table where each row contains a bookmark.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name (bookmarked)");
        builder.EndRow();

        // Data row (will be repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();

        // Bookmark tag: the bookmark name comes from the data source,
        // and the visible content is the item's name.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");

        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("Report.docx");
    }
}

// Wrapper class for the data source.
public class ReportModel
{
    public List<RowItem> Items { get; set; } = new();
}

// Represents a single row in the table.
public class RowItem
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
