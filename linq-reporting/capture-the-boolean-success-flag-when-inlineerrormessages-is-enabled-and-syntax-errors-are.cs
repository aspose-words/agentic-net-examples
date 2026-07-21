using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "template.docx");
            string outputPath = Path.Combine(workDir, "output.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the model's Name value.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Intentional syntax error – missing closing ">>".
            // This will cause the reporting engine to generate an inline error message.
            builder.Writeln("This line has a syntax error: <<[model.Name]");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Create a model instance with sample data.
            ReportModel model = new ReportModel { Name = "World" };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine to inline error messages.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. Capture the success flag; if an exception occurs,
            // treat it as a failure (success = false).
            bool success;
            try
            {
                success = engine.BuildReport(doc, model, "model");
            }
            catch (Exception)
            {
                // When a non‑recoverable syntax error occurs, BuildReport may throw.
                // In this scenario we consider the build unsuccessful.
                success = false;
            }

            // Save the generated report.
            doc.Save(outputPath);

            // Output the success flag to the console.
            Console.WriteLine($"Report build success: {success}");
            Console.WriteLine($"Output document saved to: {outputPath}");
        }
    }
}
