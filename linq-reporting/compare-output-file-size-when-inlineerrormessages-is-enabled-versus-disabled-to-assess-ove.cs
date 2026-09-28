using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessageSizeComparison
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
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document with a valid tag and an intentional error tag.
            const string templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Generate report without inline error messages.
            long sizeWithoutInline = GenerateReport(templatePath, "report_without_inline.docx", enableInline: false);

            // Generate report with inline error messages.
            long sizeWithInline = GenerateReport(templatePath, "report_with_inline.docx", enableInline: true);

            // Output file sizes.
            Console.WriteLine($"Report size without InlineErrorMessages: {sizeWithoutInline} bytes");
            Console.WriteLine($"Report size with InlineErrorMessages: {sizeWithInline} bytes");
        }

        // Creates a simple Word document containing LINQ Reporting tags.
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Valid tag – will be replaced with the Name property.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Invalid tag – property does not exist, used to demonstrate inline error messages.
            builder.Writeln("Missing property: <<[model.Missing]>>");

            doc.Save(path);
        }

        // Loads the template, builds the report with or without inline error messages, and returns the file size.
        private static long GenerateReport(string templatePath, string outputPath, bool enableInline)
        {
            // Load the template document.
            var doc = new Document(templatePath);

            // Prepare the data model.
            var model = new ReportModel();

            // Configure the reporting engine.
            var engine = new ReportingEngine();
            engine.Options = enableInline ? ReportBuildOptions.InlineErrorMessages : ReportBuildOptions.None;

            // Build the report. When inline error messages are disabled the engine will throw
            // an exception for the missing property. We catch it so the example can continue.
            try
            {
                bool success = engine.BuildReport(doc, model, "model");
                // success is true when no errors occurred; we do not need it further.
            }
            catch (Exception ex)
            {
                // Expected when InlineErrorMessages is disabled and the template contains an invalid tag.
                // The document remains in its original (template) state.
                Console.WriteLine($"BuildReport exception (expected when inline disabled): {ex.Message}");
            }

            // Save the generated document.
            doc.Save(outputPath);

            // Return the size of the generated file.
            return new FileInfo(outputPath).Length;
        }
    }
}
