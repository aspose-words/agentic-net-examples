using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBookmarkExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Bookmark name must be a non‑empty string.
        public string BookmarkName { get; set; } = "MyBookmark";

        // Sample content that will appear inside the bookmark.
        public string Title { get; set; } = "Hello World from LINQ Reporting!";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a simple template document with a bookmark tag.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert the bookmark opening tag with an expression that returns a non‑empty string.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content that will be bookmarked.
            builder.Writeln("<<[model.Title]>>");
            // Close the bookmark.
            builder.Writeln("<</bookmark>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Load the template back (simulates a real scenario where the template is read from storage).
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string outputPath = "output.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
