using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Title = "Chapter 1: Introduction",
            BookmarkName = "IntroBookmark",
            LinkText = "Go to Introduction"
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // First paragraph with a bookmark that encloses the title.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");

        builder.Writeln(); // Empty line.

        // Some placeholder content.
        builder.Writeln("This is some introductory text that will appear before the link.");

        builder.Writeln(); // Empty line.

        // Hyperlink that points to the bookmark defined above.
        builder.Writeln("<<link [model.BookmarkName] [model.LinkText]>>");

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("InternalLinkReport.docx");
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
