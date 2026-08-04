using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Bookmark name may be empty; initialize to empty string to avoid nullable warnings.
    public string BookmarkName { get; set; } = string.Empty;
    public string Title { get; set; } = "Sample Title";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            // Intentionally leave BookmarkName empty to trigger the warning.
            BookmarkName = string.Empty,
            Title = "Hello from LINQ Reporting"
        };

        // Log a warning if the bookmark name expression would be empty.
        if (string.IsNullOrEmpty(model.BookmarkName))
        {
            Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
        }

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Conditional block: include the bookmark only when the name is not empty.
        builder.Writeln("<<if [model.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        // Remove empty paragraphs that may remain after the conditional block is skipped.
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
