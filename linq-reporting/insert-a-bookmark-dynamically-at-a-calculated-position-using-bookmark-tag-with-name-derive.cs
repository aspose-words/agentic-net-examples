using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Title = "First Item" },
                new Item { Id = 2, Title = "Second Item" },
                new Item { Id = 3, Title = "Third Item" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report generated with dynamic bookmarks:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the root object named "model".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithBookmarks.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model with a calculated bookmark name.
public class Item
{
    public int Id { get; set; }
    public string Title { get; set; } = "";
    public string BookmarkName => $"Item_{Id}";
}
