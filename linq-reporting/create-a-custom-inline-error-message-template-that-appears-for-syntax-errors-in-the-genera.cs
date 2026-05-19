using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingInlineError
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public string Content { get; set; } = "This is a valid content.";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tags.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<[model.Content]>>");

            // Intentionally malformed tag – missing closing ">>".
            // This will trigger a syntax error that the ReportingEngine can inline.
            builder.Writeln("<<[model.Title]"); // <-- syntax error

            // Save the template (required before building the report).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine
            {
                // Enable inline error messages so that syntax errors appear in the output document.
                Options = ReportBuildOptions.InlineErrorMessages
            };

            bool success;
            try
            {
                // BuildReport returns true if parsing succeeded without errors.
                success = engine.BuildReport(doc, model, "model");
            }
            catch (Exception ex)
            {
                // If an unexpected exception occurs, treat the build as failed.
                Console.WriteLine($"Exception during report generation: {ex.Message}");
                success = false;
            }

            // Save the generated document.
            doc.Save(resultPath);

            // Output the result.
            Console.WriteLine($"Report build success: {success}");
            Console.WriteLine($"Result saved to: {resultPath}");
        }
    }
}
