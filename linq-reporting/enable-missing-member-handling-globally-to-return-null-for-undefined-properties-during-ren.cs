using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a defined property only.
    public class Model
    {
        public string Name { get; set; } = "John Doe";
        // Note: No property named MissingProperty – it will be missing in the template.
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Write a line that references an existing property.
            builder.Writeln("Name: <<[model.Name]>>");

            // Write a line that references a missing property.
            // With AllowMissingMembers this will be treated as null (empty output).
            builder.Writeln("Missing: <<[model.MissingProperty]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // 3. Prepare the data source.
            var data = new Model();

            // 4. Configure the ReportingEngine to allow missing members.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                // Optional: customize the message shown for a missing plain reference.
                MissingMemberMessage = string.Empty
            };

            // 5. Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(reportDoc, data, "model");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
