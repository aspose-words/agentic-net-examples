using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Bookmark name must be a non‑empty string.
    public string BookmarkName { get; set; } = "SampleBookmark";

    // Sample content to place inside the bookmark.
    public string Title { get; set; } = "Hello World";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting bookmark tag whose name comes from the model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        // Content that will be inside the bookmark.
        builder.Writeln("<<[model.Title]>>");
        // Close the bookmark tag.
        builder.Writeln("<</bookmark>>");

        // Build the report using the model as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, new ReportModel(), "model");

        // Save the generated document.
        doc.Save("ReportWithBookmark.docx");
    }
}
