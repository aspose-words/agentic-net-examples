using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new ReportModel
        {
            Title = "Dynamic Bookmark Example",
            // Bookmark name derived from data fields (e.g., title and timestamp).
            BookmarkName = $"BM_{DateTime.Now:yyyyMMdd_HHmmss}"
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Insert a bookmark tag with a name derived from the model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("This paragraph is inside the dynamically created bookmark.");
        builder.Writeln("<</bookmark>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Ensure output directory exists.
        var outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Save the generated document.
        var outputPath = Path.Combine(outputDir, "result.docx");
        doc.Save(outputPath);
    }
}

// Public data model class with required properties.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
