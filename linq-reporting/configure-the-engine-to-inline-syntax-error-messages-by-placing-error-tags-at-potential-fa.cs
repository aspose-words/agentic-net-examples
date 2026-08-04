using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Non‑nullable properties are initialized to avoid warnings.
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Valid tags.
            builder.Writeln("Customer Name: <<[model.Name]>>");
            builder.Writeln("Customer Age: <<[model.Age]>>");

            // Tag that references a missing member – will produce an inline error.
            builder.Writeln("Missing Property: <<[model.Missing]>>");

            // Malformed tag – also triggers an inline error.
            builder.Writeln("Malformed Tag: <<[model.Age>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a separate load step).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Name = "John Doe",
                Age = 30
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine to inline error messages.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The boolean indicates whether parsing succeeded.
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);

            // Output the result status.
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
