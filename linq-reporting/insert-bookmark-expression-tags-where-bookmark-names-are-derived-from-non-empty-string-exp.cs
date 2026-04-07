using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkTagExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Name of the bookmark – must be a non‑empty string.
        public string BookmarkName { get; set; } = "MyBookmark";

        // Text that will appear inside the bookmark.
        public string Title { get; set; } = "Hello from Aspose.Words!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a bookmark tag whose name is taken from the model.
            // The expression [model.BookmarkName] must evaluate to a non‑empty string.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content inside the bookmark – another expression.
            builder.Writeln("<<[model.Title]>>");
            // Closing tag for the bookmark.
            builder.Writeln("<</bookmark>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data source.
            // -------------------------------------------------
            var model = new ReportModel
            {
                BookmarkName = "MyBookmark",          // Non‑empty bookmark name.
                Title = "Hello from Aspose.Words!"    // Sample content.
            };

            // -------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            var engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
