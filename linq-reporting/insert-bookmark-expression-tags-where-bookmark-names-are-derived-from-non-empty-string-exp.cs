using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Bookmark name expression – must be a non‑empty string.
        public string BookmarkName { get; set; } = "MyBookmark";

        // Content that will appear inside the bookmark.
        public string Title { get; set; } = "Hello from LINQ Reporting!";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document with a bookmark tag.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a bookmark tag whose name is taken from the model expression.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content that will be bookmarked.
            builder.Writeln("<<[model.Title]>>");
            // Close the bookmark tag.
            builder.Writeln("<</bookmark>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var model = new ReportModel();

            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
