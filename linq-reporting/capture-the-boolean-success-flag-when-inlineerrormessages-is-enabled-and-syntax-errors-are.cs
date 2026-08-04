using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; }

        public ReportModel(string name) => Name = name;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "ReportOutput.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the model's Name value.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Incorrect tag – missing closing ">>" to trigger a syntax error.
            // This will cause the reporting engine to generate an inline error message
            // when the InlineErrorMessages option is enabled.
            builder.Writeln("This line contains a syntax error <<[model.Name]");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the reporting engine with InlineErrorMessages enabled.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Sample data source.
            ReportModel model = new ReportModel("World");

            // Build the report and capture the success flag.
            bool success;
            try
            {
                // The overload that includes the data source name allows the template
                // to reference the root object via <<[model.Name]>>.
                success = engine.BuildReport(reportDoc, model, "model");
            }
            catch (Exception ex)
            {
                // If an unexpected exception occurs, treat the build as unsuccessful.
                Console.WriteLine($"Exception during report build: {ex.Message}");
                success = false;
            }

            // Output the success flag to the console.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
