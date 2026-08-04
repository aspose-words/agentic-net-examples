using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorReportingExample
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Correct tag – will be replaced with the model's Name value.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Intentionally malformed tag – extra '>' creates a syntax error.
            // With InlineErrorMessages enabled the engine will insert an error message in place of this tag.
            builder.Writeln("This line contains a bad tag: <<[model.Name]>>>>>");

            // Prepare the data source.
            Model model = new Model();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned flag indicates whether parsing succeeded.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            string outputPath = "ReportWithInlineErrors.docx";
            doc.Save(outputPath);

            // Output the result status – useful for debugging but not required for the example.
            Console.WriteLine($"Report built successfully: {success}");
            Console.WriteLine($"Output saved to: {outputPath}");
        }
    }
}
