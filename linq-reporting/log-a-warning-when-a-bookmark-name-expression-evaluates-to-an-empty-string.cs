using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Bookmark name that will be used in the template.
    // Intentionally left empty to trigger the warning and the conditional tag.
    public string BookmarkName { get; set; } = string.Empty;

    // Some content to place inside the bookmark.
    public string Title { get; set; } = "Sample Title";
}

public class Program
{
    public static void Main()
    {
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Use an IF tag so the bookmark block is only processed when the name is not empty.
        // The condition must evaluate to a Boolean value.
        builder.Writeln("<<if [model.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real‑world scenario).
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            // BookmarkName left empty to demonstrate the warning.
            BookmarkName = string.Empty,
            Title = "Hello from LINQ Reporting"
        };

        // -----------------------------------------------------------------
        // 4. Log a warning if the bookmark name evaluates to an empty string.
        // -----------------------------------------------------------------
        if (string.IsNullOrWhiteSpace(model.BookmarkName))
        {
            Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
        }

        // -----------------------------------------------------------------
        // 5. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
