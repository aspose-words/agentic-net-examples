using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBookmarkExample
{
    // Simple data model used by the LINQ Reporting engine.
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
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a bookmark tag whose name comes from the model's BookmarkName property.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content that will be placed inside the bookmark.
            builder.Writeln("<<[model.Title]>>");
            // Close the bookmark tag.
            builder.Writeln("<</bookmark>>");

            // Save the template to disk before building the report.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportModel model = new ReportModel(); // Sample data.

            ReportingEngine engine = new ReportingEngine();
            // BuildReport overload that specifies the root name ("model").
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(resultPath);

            // Inform the user where files are written (optional, not interactive).
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated at: {resultPath}");
        }
    }
}
