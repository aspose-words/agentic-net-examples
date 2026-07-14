using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Valid string property.
        public string Name { get; set; } = string.Empty;

        // Valid integer property.
        public int Age { get; set; }

        // This property is intentionally omitted to trigger a conversion error in the template.
        // The engine will insert an inline error message because ReportBuildOptions.InlineErrorMessages is enabled.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a Word template programmatically and insert LINQ
            // Reporting tags. One of the tags references a non‑existent member to
            // demonstrate inline error messages.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");
            // This tag will cause a conversion/missing‑member error.
            builder.Writeln("Missing: <<[model.NonExisting]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report using the data model.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Populate the model with sample data.
            ReportModel model = new ReportModel
            {
                Name = "John Doe",
                Age = 30
            };

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The third parameter ("model") matches the root name used
            // in the template tags (<<[model.Property]>>).
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report generation successful: {success}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
