using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model with only the Name property.
    public class Model
    {
        public string Name { get; set; } = string.Empty;
        // Age is intentionally omitted to demonstrate AllowMissingMembers.
    }

    public class Program
    {
        public static void Main()
        {
            // File paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a Word template that contains LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Correct tag – will be replaced with the Name value.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Missing member tag – Age does not exist in Model.
            // With AllowMissingMembers this will be treated as null (empty output).
            builder.Writeln("Customer Age: <<[model.Age]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare the data source.
            // -----------------------------------------------------------------
            Model data = new Model { Name = "John Doe" };

            // -----------------------------------------------------------------
            // Step 4: Configure and run the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned bool indicates success (true) or failure (false).
            bool success = engine.BuildReport(reportDoc, data, "model");

            // -----------------------------------------------------------------
            // Step 5: Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            // Indicate the result (no interactive prompts).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
