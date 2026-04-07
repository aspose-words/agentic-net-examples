using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Bookmark name that will be used in the <<bookmark>> tag.
        // Initialized to an empty string to demonstrate the warning scenario.
        public string BookmarkName { get; set; } = string.Empty;

        // Additional content that will appear inside the bookmark.
        public string Content { get; set; } = "Sample bookmarked text.";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains a conditional bookmark tag.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Use an <<if>> block so the bookmark tag is only emitted when the name is not empty.
            // This prevents the engine from throwing an exception for an empty bookmark name.
            builder.Writeln("<<if [model.BookmarkName]>>");
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Content]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</if>>");

            // Save the template to disk (required by the lifecycle rule).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk before building the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Intentionally leave BookmarkName empty to trigger the warning.
                BookmarkName = string.Empty,
                Content = "This is the content inside the bookmark."
            };

            // -----------------------------------------------------------------
            // 4. Log a warning if the bookmark name evaluates to an empty string.
            // -----------------------------------------------------------------
            if (string.IsNullOrEmpty(model.BookmarkName))
            {
                Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
            }

            // -----------------------------------------------------------------
            // 5. Build the report using the ReportingEngine.
            //    Use InlineErrorMessages to prevent an exception when the bookmark name is empty.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            bool success = engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated and saved to '{outputPath}'. Success flag: {success}");
        }
    }
}
