using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
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

            // -----------------------------------------------------------------
            // Step 1: Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the model's Name value.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Intentional syntax error – missing closing brackets.
            // This will trigger an inline error message when InlineErrorMessages is enabled.
            builder.Writeln("Broken Tag: <<[model.Name]");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var model = new Model();

            var engine = new ReportingEngine
            {
                // Enable inline error messages.
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // BuildReport returns a Boolean indicating whether parsing succeeded.
            bool success;
            try
            {
                success = engine.BuildReport(doc, model, "model");
            }
            catch (Exception ex)
            {
                // If an unexpected exception occurs, treat the build as unsuccessful.
                Console.WriteLine($"Exception during report build: {ex.Message}");
                success = false;
            }

            // Save the generated report.
            doc.Save(outputPath);

            // Output the success flag.
            Console.WriteLine($"Report build success: {success}");
        }
    }
}
