using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Text that will appear inside the bookmark.
        public string Title { get; set; } = string.Empty;

        // Name of the bookmark. Must be a valid identifier.
        public string BookmarkName { get; set; } = string.Empty;

        // Text displayed for the hyperlink that points to the bookmark.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple introductory paragraph.
            builder.Writeln("Demo: Internal document link using bookmarks and hyperlinks.");

            // Define a bookmark whose name comes from the data model.
            // The bookmark encloses the title text.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");

            // Paragraph that contains a hyperlink pointing to the bookmark.
            // The first expression is the bookmark name, the second is the link display text.
            builder.Writeln("Navigate to the section: <<link [model.BookmarkName] [model.LinkText]>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            var model = new ReportModel
            {
                Title = "Target Section",
                BookmarkName = "MyBookmark",
                LinkText = "Click Here"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };

            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);

            // Simple console output to indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
