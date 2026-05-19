using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with only a Name property.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create an output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Step 1: Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Normal field – will be replaced with the model's Name value.
            builder.Writeln("Name: <<[model.Name]>>");

            // Missing field – does not exist in the model. With AllowMissingMembers,
            // it will be treated as null and the MissingMemberMessage will be inserted.
            builder.Writeln("Missing: <<[model.MissingField]>>");

            // Syntax error – an unsupported switch is used.
            // With InlineErrorMessages, the engine will embed an error message directly into the output.
            builder.Writeln("Syntax error: <<[model.Name] -unknown>>");

            // Save the template to disk.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template document for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel { Name = "John Doe" };

            // -----------------------------------------------------------------
            // Step 4: Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;
            engine.MissingMemberMessage = "N/A";

            // Build the report. The third argument is the name used in the template tags.
            bool success = engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // Step 5: Save the generated report.
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);

            // Output the result status.
            Console.WriteLine($"Report generation successful: {success}");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
