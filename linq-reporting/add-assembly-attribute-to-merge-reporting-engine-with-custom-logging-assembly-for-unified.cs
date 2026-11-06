using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any legacy encodings that might be needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create a Word document template programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the root data object.
            ReportModel model = new ReportModel { Name = "Aspose.Words" };

            // Configure the reporting engine.
            ReportingEngine.UseReflectionOptimization = true;
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report.
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Output simple status (no interactive prompts).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{reportPath}'.");
        }
    }
}
