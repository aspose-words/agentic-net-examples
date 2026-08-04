using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Model class exposed to the template.
    public class ReportModel
    {
        // Initialized to avoid nullable warnings.
        public DateTime CurrentUtcNow { get; set; } = DateTime.UtcNow;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the working directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a tag that references the model's CurrentUtcNow property.
            builder.Writeln("Current UTC time: <<[model.CurrentUtcNow]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create and configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.

            // Build the report using the root object name "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            loadedTemplate.Save(reportPath);

            // Inform the user where the files are located (no interactive input required).
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to:   {reportPath}");
        }
    }
}
