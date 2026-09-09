using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class Model
    {
        // Initialized to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
        // No property named 'Missing' – this will be accessed only via the template.
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write a line that references an existing property.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Write a line that references a missing property.
            // With AllowMissingMembers option this will be treated as null (empty).
            builder.Writeln("Missing Property: <<[model.Missing]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Create the data source.
            var model = new Model();

            // Configure the reporting engine to treat missing members as null.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                // Optional: customize the message printed for a plain missing member reference.
                MissingMemberMessage = string.Empty
            };

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
        }
    }
}
