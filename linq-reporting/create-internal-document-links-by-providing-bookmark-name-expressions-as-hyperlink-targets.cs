using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Title displayed inside the bookmark.
        public string Title { get; set; } = string.Empty;

        // Name of the bookmark that will be created.
        public string BookmarkName { get; set; } = string.Empty;

        // Text shown for the hyperlink that points to the bookmark.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Title = "Section 1 – Introduction",
                BookmarkName = "Section1",
                LinkText = "Jump to Section 1"
            };

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Define a bookmark whose name comes from the model.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            // Content inside the bookmark.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");

            builder.Writeln(); // Empty paragraph.

            // Hyperlink that points to the same bookmark.
            builder.Writeln("See details: <<link [model.BookmarkName] [model.LinkText]>>");

            // Save the template to disk (required before BuildReport according to rules).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var document = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root object name is "model".
            engine.BuildReport(document, model, "model");

            // Save the final document.
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
