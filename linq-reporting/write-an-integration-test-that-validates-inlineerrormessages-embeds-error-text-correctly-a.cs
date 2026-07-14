using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesTest
{
    // Simple data model – does not contain the missingObject property used in the template.
    public class Model
    {
        // No members needed for this test.
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // ---------- Create the template document ----------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // This tag references a non‑existent member and will cause a parsing error.
            builder.Writeln("<<[missingObject.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------- Load the template and build the report ----------
            Document reportDoc = new Document(templatePath);

            // Configure the reporting engine to embed inline error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report using a model that does not provide the missing member.
            bool success = engine.BuildReport(reportDoc, new Model(), "model");

            // Save the generated report.
            reportDoc.Save(resultPath);

            // ---------- Verify the outcome ----------
            // The document text should contain an error message inserted by the engine.
            string docText = reportDoc.GetText();

            bool containsError = docText.Contains("Error", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"BuildReport success flag: {success}");
            Console.WriteLine($"Document contains inline error message: {containsError}");
        }
    }
}
