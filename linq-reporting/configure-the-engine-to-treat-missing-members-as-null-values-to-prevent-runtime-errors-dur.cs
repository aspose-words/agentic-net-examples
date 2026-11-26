using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingMissingMembers
{
    // Sample data model.
    public class Model
    {
        // Non‑nullable property initialized to avoid warnings.
        public string Name { get; set; } = "John Doe";

        // Property that does not exist in the template – will be missing.
        // No definition needed; the engine will treat it as null.
    }

    public class Program
    {
        public static void Main()
        {
            // Required for some Aspose.Words features (e.g., code page support).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new Model();

            // -----------------------------------------------------------------
            // Create a template document programmatically.
            // -----------------------------------------------------------------
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags.
            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln("Missing Member (should be empty): <<[model.Missing]>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // Load the template for reporting.
            var template = new Document(templatePath);

            // -----------------------------------------------------------------
            // Configure the reporting engine.
            // Use InlineErrorMessages to avoid runtime exceptions for missing members.
            // This will render missing members as empty strings (no error thrown).
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(template, model, "model");

            // Save the generated report.
            var outputDir = "output";
            Directory.CreateDirectory(outputDir);
            var outputPath = Path.Combine(outputDir, "report.docx");
            template.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
