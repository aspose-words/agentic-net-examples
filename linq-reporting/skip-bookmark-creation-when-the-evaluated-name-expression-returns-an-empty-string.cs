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
                new Item { Title = "First Section", BookmarkName = "FirstSec" },
                new Item { Title = "Second Section", BookmarkName = "" }, // No bookmark.
                new Item { Title = "Third Section", BookmarkName = "ThirdSec" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // If the bookmark name is not empty, create a bookmark around the title.
        builder.Writeln("<<if [item.BookmarkName != \"\"]>><<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>><</if>>");

        // If the bookmark name is empty, just write the title without a bookmark.
        builder.Writeln("<<if [item.BookmarkName == \"\"]>><<[item.Title]>>");
        builder.Writeln("<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine
        {
            // Remove empty paragraphs that may appear after tags are omitted.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("ReportWithConditionalBookmarks.docx");
    }
}

// Root data model referenced by the template.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item used in the foreach loop.
public class Item
{
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
