using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesTest
{
    // Simple data model used as the root object for the report.
    public class Model
    {
        // Existing property that will be rendered correctly.
        public string Name { get; set; } = "Sample Name";

        // No property named MissingProperty – this will trigger an inline error.
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare file paths.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a correct tag.
            builder.Writeln("Name: <<[model.Name]>>");
            // Write a tag that references a non‑existent property – this will cause an error.
            builder.Writeln("Missing: <<[model.MissingProperty]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for reporting.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Build the report with InlineErrorMessages enabled.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            Model model = new Model();

            bool success = engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 4. Verify the result.
            // -------------------------------------------------
            string reportText = reportDoc.GetText();

            bool containsError = reportText.Contains("Error", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"BuildReport success flag: {success}");
            Console.WriteLine($"Document contains inline error message: {containsError}");

            // Save the generated document for manual inspection if needed.
            reportDoc.Save(outputPath);
        }
    }
}
