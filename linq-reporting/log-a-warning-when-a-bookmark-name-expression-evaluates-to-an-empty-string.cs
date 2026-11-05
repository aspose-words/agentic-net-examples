using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Document with a conditional bookmark generated via LINQ Reporting:");

        // Conditional bookmark tag – will be processed only when the name is not empty.
        builder.Writeln("<<if [model.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("This is the bookmarked content.");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the data model.
        // -------------------------------------------------
        var model = new ReportModel
        {
            // Intentionally set to empty to trigger the warning.
            BookmarkName = ""
        };

        // Log a warning if the bookmark name evaluates to an empty string.
        if (string.IsNullOrEmpty(model.BookmarkName))
        {
            Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
        }

        // -------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var engine = new ReportingEngine
        {
            // Inline error messages will be inserted into the document instead of throwing.
            Options = ReportBuildOptions.InlineErrorMessages
        };
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// Public data model class used by the template.
public class ReportModel
{
    // Bookmark name used in the <<bookmark>> tag.
    public string BookmarkName { get; set; } = string.Empty;

    // Additional sample property (not used in this example but kept for completeness).
    public string Title { get; set; } = "Sample Title";
}
