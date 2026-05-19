using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialized to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = "Output";
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // -----------------------------------------------------------------
            // Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Valid tag – will be replaced with the model's Name property.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Deliberate syntax error – references a non‑existent member.
            // With InlineErrorMessages enabled, the engine will insert an error message here.
            builder.Writeln("Invalid tag: <<[model.NonExistent]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // Configure the engine to inline error messages.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report using the model as the data source.
            bool success = engine.BuildReport(doc, new Model(), "model");

            // Output the success flag (true if parsing succeeded, false otherwise).
            Console.WriteLine($"BuildReport success: {success}");

            // Save the generated report.
            doc.Save(resultPath);
        }
    }
}
