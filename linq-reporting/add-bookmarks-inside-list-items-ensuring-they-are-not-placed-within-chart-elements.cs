using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with Bookmarks in List Items");
        builder.Writeln("<<foreach [item in Items]>>");
        // Bookmark tag placed inside the list item (paragraph).
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First Item", BookmarkName = "bm1" },
                new Item { Title = "Second Item", BookmarkName = "bm2" },
                new Item { Title = "Third Item", BookmarkName = "bm3" }
            }
        };

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportWithBookmarks.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
