using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");
            string errorLogPath = Path.Combine(outputDir, "error.log");

            // -----------------------------------------------------------------
            // 1. Create a template document with both a valid and an invalid tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Valid expression tag.
            builder.Writeln("<<[model.Name]>>");

            // Invalid expression tag (missing closing ">>").
            builder.Writeln("<<[model.Name]");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the reporting engine without InlineErrorMessages.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Ensure InlineErrorMessages flag is NOT set.
                Options = ReportBuildOptions.None
            };

            // -----------------------------------------------------------------
            // 4. Build the report and capture any syntax errors.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            try
            {
                // BuildReport will throw an exception because the template contains a syntax error.
                bool success = engine.BuildReport(doc, model, "model");

                // If, for any reason, the build succeeds, save the generated report.
                if (success)
                {
                    doc.Save(reportPath);
                }
            }
            catch (Exception ex)
            {
                // Log the exception message (syntax error details) to a file.
                File.WriteAllText(errorLogPath, ex.Message);
            }
        }
    }
}
