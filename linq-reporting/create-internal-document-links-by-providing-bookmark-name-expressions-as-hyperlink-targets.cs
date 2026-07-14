using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Define a bookmark whose name comes from the data source.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        // The content of the bookmark – the title of the item.
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Insert a hyperlink that points to the same bookmark.
        // The display text is "Go to " followed by the title.
        builder.Writeln("<<link [item.BookmarkName] [\"Go to \" + item.Title]>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new() { Title = "Introduction", BookmarkName = "bmIntro" },
                new() { Title = "Chapter 1", BookmarkName = "bmChapter1" },
                new() { Title = "Conclusion", BookmarkName = "bmConclusion" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item containing a title and a bookmark name.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
