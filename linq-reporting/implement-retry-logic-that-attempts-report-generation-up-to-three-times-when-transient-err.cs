using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingRetry
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Title { get; set; } = "Sample Report";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Report Title: <<[model.Title]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed.

            const int maxAttempts = 3;
            bool success = false;

            // -------------------------------------------------
            // Retry logic: attempt to build the report up to three times.
            // -------------------------------------------------
            for (int attempt = 1; attempt <= maxAttempts && !success; attempt++)
            {
                try
                {
                    // BuildReport overload that allows referencing the root object name.
                    success = engine.BuildReport(reportDoc, model, "model");
                }
                catch (Exception ex)
                {
                    // In a real scenario, inspect the exception to determine if it is transient.
                    // For this example, any exception triggers a retry until the max attempts are reached.
                    Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");

                    if (attempt == maxAttempts)
                    {
                        // Rethrow the exception after the final attempt.
                        throw;
                    }

                    // Optionally, introduce a short delay before retrying.
                    System.Threading.Thread.Sleep(500);
                }
            }

            // -------------------------------------------------
            // Save the generated report if successful.
            // -------------------------------------------------
            if (success)
            {
                reportDoc.Save(outputPath);
                Console.WriteLine($"Report generated successfully and saved to '{outputPath}'.");
            }
        }
    }
}
