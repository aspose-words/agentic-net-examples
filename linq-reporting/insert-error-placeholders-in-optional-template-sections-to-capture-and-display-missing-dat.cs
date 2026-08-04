using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingErrorPlaceholders
{
    // Simple data model for the report.
    public class ReportModel
    {
        // Initialized to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        // Note: No property named 'NonExistent' – this will trigger a missing‑member warning.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags. The <<error>> tag will display any inline error messages.
            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln("<<error>>"); // Placeholder for potential errors on the previous line.

            builder.Writeln("Customer Age: <<[model.Age]>>");
            builder.Writeln("<<error>>"); // Placeholder for potential errors on the previous line.

            // This field does not exist in the model and will generate a missing‑member warning.
            builder.Writeln("Missing Field: <<[model.NonExistent]>>");
            builder.Writeln("<<error>>"); // Capture the warning for the missing field.

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario where the template is stored).
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel(); // All required properties are initialized.

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine with InlineErrorMessages enabled.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // BuildReport returns a bool indicating success when InlineErrorMessages is set.
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Output the success flag to the console (no interactive prompts).
            Console.WriteLine($"Report generation successful: {success}");
            Console.WriteLine($"Template saved to: {Path.GetFullPath(templatePath)}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
