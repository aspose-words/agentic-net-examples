using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Bookmark name that will be evaluated in the template.
        // Initialized to an empty string to demonstrate the warning scenario.
        public string BookmarkName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document with a bookmark tag.
            const string templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("This is bookmarked content.");
            builder.Writeln("<</bookmark>>");
            builder.Document.Save(templatePath);

            // Step 2: Load the template document.
            var doc = new Document(templatePath);

            // Step 3: Prepare the data model.
            var model = new ReportModel(); // BookmarkName is empty.

            // Step 4: Log a warning and provide a fallback bookmark name if the expression evaluates to an empty string.
            if (string.IsNullOrEmpty(model.BookmarkName))
            {
                Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
                // Provide a non‑empty placeholder name so the reporting engine does not throw.
                model.BookmarkName = "DefaultBookmark";
            }

            // Step 5: Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 6: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
