using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a bookmark tag that uses a model expression for the bookmark name.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            BookmarkName = "", // Intentionally empty to trigger the warning.
            Title = "Sample Report Title"
        };

        // Log a warning if the bookmark name evaluates to an empty string and provide a fallback name.
        if (string.IsNullOrEmpty(model.BookmarkName))
        {
            Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
            model.BookmarkName = "DefaultBookmark";
        }

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "output.docx");
        doc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }

    // Public data model class required by the LINQ Reporting engine.
    public class ReportModel
    {
        public string BookmarkName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
    }
}
