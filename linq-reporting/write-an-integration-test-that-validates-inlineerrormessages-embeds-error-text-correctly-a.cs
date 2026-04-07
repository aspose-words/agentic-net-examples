using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesDemo
{
    // Simple data model with one valid property.
    public class ReportModel
    {
        public string ExistingProperty { get; set; } = "This is a valid value.";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // This tag references a missing property and will trigger an error.
            builder.Writeln("<<[model.MissingProperty]>>");

            // This tag references an existing property and will be replaced normally.
            builder.Writeln("<<[model.ExistingProperty]>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 3. Build the report with InlineErrorMessages enabled.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // BuildReport returns a flag indicating whether parsing succeeded.
            bool success = engine.BuildReport(template, model, "model");

            // Save the resulting document for inspection.
            string resultPath = Path.Combine(outputDir, "Report.docx");
            template.Save(resultPath);

            // -----------------------------------------------------------------
            // 4. Output the test results.
            // -----------------------------------------------------------------
            Console.WriteLine($"BuildReport success flag: {success}");
            Console.WriteLine("Document content:");
            Console.WriteLine(template.GetText());

            // Expected: success == false because of the missing property,
            // and the document text should contain an inline error message.
        }
    }
}
