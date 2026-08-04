using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Name of the bookmark that will be created in the document.
        public string BookmarkName { get; set; } = "MyBookmark";

        // Text that will appear inside the bookmark.
        public string Title { get; set; } = "This is the bookmarked content.";

        // Text displayed for the hyperlink that points to the bookmark.
        public string LinkText { get; set; } = "Go to bookmark";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the final report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Introductory paragraph.
            builder.Writeln("Document demonstrating internal links with LINQ Reporting.");

            // Bookmark tag: creates a bookmark whose name is taken from the model.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content inside the bookmark.
            builder.Writeln("<<[model.Title]>>");
            // Closing tag for the bookmark.
            builder.Writeln("<</bookmark>>");

            // Add an empty paragraph for visual separation.
            builder.Writeln();

            // Link tag: creates a hyperlink that points to the bookmark name from the model.
            // The first expression is the bookmark name, the second is the display text.
            builder.Writeln("<<link [model.BookmarkName] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            // Build the report using the model; the root object name is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
