using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBookmarkExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Title that will appear inside the bookmark (if created).
        public string Title { get; set; } = "Sample Report Title";

        // Name of the bookmark. When empty, the bookmark should be omitted.
        public string BookmarkName { get; set; } = "";
    }

    class Program
    {
        static void Main()
        {
            // Folder for generated files.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the resulting reports.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportWithBookmarkPath = Path.Combine(outputDir, "Report_WithBookmark.docx");
            string reportWithoutBookmarkPath = Path.Combine(outputDir, "Report_WithoutBookmark.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Conditional block: create a bookmark only when BookmarkName is not empty.
            builder.Writeln("<<if [model.BookmarkName != \"\"]>>");
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build a report where the bookmark is created.
            // -----------------------------------------------------------------
            Document docWithBookmark = new Document(templatePath);
            var modelWithBookmark = new ReportModel
            {
                Title = "Report with Bookmark",
                BookmarkName = "MyBookmark"
            };

            var engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear when the condition is false.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            engine.BuildReport(docWithBookmark, modelWithBookmark, "model");
            docWithBookmark.Save(reportWithBookmarkPath);

            // -----------------------------------------------------------------
            // 3. Load the template and build a report where the bookmark is skipped.
            // -----------------------------------------------------------------
            Document docWithoutBookmark = new Document(templatePath);
            var modelWithoutBookmark = new ReportModel
            {
                Title = "Report without Bookmark",
                BookmarkName = "" // Empty name -> bookmark should not be created.
            };

            engine.BuildReport(docWithoutBookmark, modelWithoutBookmark, "model");
            docWithoutBookmark.Save(reportWithoutBookmarkPath);

            // Indicate completion.
            Console.WriteLine("Reports generated:");
            Console.WriteLine(reportWithBookmarkPath);
            Console.WriteLine(reportWithoutBookmarkPath);
        }
    }
}
