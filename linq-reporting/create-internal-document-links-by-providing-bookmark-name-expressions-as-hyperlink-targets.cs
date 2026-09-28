using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace InternalLinkExample
{
    // Data model for the report.
    public class ReportModel
    {
        public string BookmarkName { get; set; } = "MyBookmark";
        public string Title { get; set; } = "Section Title";
        public string LinkText { get; set; } = "Go to Section";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for template and result.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Intro text.
            builder.Writeln("Document demonstrating internal links using bookmarks.");

            // Bookmark definition with expressions.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");

            // Some spacing.
            builder.Writeln();

            // Hyperlink that points to the bookmark defined above.
            builder.Writeln("<<link [model.BookmarkName] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save(resultPath);
        }
    }
}
