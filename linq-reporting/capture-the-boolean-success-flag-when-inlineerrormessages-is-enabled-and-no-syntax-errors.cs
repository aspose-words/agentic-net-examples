using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingInlineErrorDemo
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
            // Prepare a temporary folder for the files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinqReportingDemo");
            Directory.CreateDirectory(workDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(workDir, "Template.docx");
            string reportPath = Path.Combine(workDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required before building the report).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine with InlineErrorMessages.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // -----------------------------------------------------------------
            // 5. Build the report and capture the success flag.
            // -----------------------------------------------------------------
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);

            // Output the result.
            Console.WriteLine($"Report generated successfully: {success}");
            Console.WriteLine($"Template path: {templatePath}");
            Console.WriteLine($"Report path: {reportPath}");
        }
    }
}
