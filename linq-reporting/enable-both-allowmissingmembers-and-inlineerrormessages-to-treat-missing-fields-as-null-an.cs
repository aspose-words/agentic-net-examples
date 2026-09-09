using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with only one property.
    public class ReportModel
    {
        public string Name { get; set; } = "John Doe";
        // Note: Age property is intentionally omitted to demonstrate AllowMissingMembers.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a valid tag that references an existing member.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Write a tag that references a missing member (Age). With AllowMissingMembers this will be treated as null.
            builder.Writeln("Customer Age: <<[model.Age]>>");

            // Write a malformed tag to trigger a syntax error. InlineErrorMessages will embed the error message.
            builder.Writeln("Malformed Tag Example: <<[model.Name] -unknownSwitch>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine with both options.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages,
                MissingMemberMessage = "Missing"
            };

            // Build the report. The overload with dataSourceName allows the template to reference the root object as "model".
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag (true indicates the template was parsed without fatal errors).
            Console.WriteLine($"Report build success: {success}");

            // Save the generated document.
            const string outputPath = "ReportWithMissingMembers.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
