using System;
using System.IO;
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
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Result.docx");

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Hello, <<[model.Name]>>!");
            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Create sample data.
            Model model = new Model();

            // -----------------------------------------------------------------
            // Configure the reporting engine.
            // -----------------------------------------------------------------
            // Disable reflection optimization for this template.
            ReportingEngine.UseReflectionOptimization = false;

            // Create the engine (ReportingEngine does not implement IDisposable, so no using block).
            ReportingEngine engine = new ReportingEngine();

            // Build the report.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(resultPath);
        }
    }
}
