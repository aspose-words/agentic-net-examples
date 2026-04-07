using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkLinqReportingExample
{
    // Data model used as the root object for the LINQ Reporting engine.
    public class ReportModel
    {
        // Title that will be displayed inside the bookmark.
        public string Title { get; set; } = "Sample Section";

        // Bookmark name derived from the Title (spaces replaced with underscores).
        public string BookmarkName { get; set; } = "Sample_Section";

        // Additional content that could be placed inside the bookmark.
        public string Content { get; set; } = "This is the bookmarked content.";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting bookmark tag.
            // The bookmark name will be taken from the data model (model.BookmarkName).
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content inside the bookmark: title and additional text.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<[model.Content]>>");
            builder.Writeln("<</bookmark>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "BookmarkTemplate.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Title = "Dynamic Section",
                // Derive the bookmark name from the title (replace spaces with underscores).
                BookmarkName = "Dynamic_Section",
                Content = "Content placed inside the dynamically created bookmark."
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(report, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "BookmarkReport.docx";
            report.Save(outputPath);
        }
    }
}
