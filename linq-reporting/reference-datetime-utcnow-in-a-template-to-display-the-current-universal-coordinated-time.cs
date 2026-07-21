using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class whose property will be referenced in the template.
    public class ReportModel
    {
        // Initialized to the current UTC time at object creation.
        public DateTime CurrentUtc { get; set; } = DateTime.UtcNow;
    }

    public class Program
    {
        public static void Main()
        {
            // Define file names for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a line that displays the current UTC time using a LINQ Reporting tag.
            builder.Writeln("Current UTC time: <<[model.CurrentUtc]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            // Load the previously saved template.
            Document reportDoc = new Document(templatePath);

            // Create the data model instance.
            ReportModel model = new ReportModel();

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);

            // Optional: Output the result text to the console for verification.
            Console.WriteLine("Report generated successfully.");
            Console.WriteLine("Report content:");
            Console.WriteLine(reportDoc.GetText());
        }
    }
}
