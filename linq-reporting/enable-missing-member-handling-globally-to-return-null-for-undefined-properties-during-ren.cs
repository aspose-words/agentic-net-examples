using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with only one defined property.
    public class ReportModel
    {
        public string Existing { get; set; } = string.Empty;
        // Note: No property named "Missing" is defined.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The template references an existing property and a missing one.
            builder.Writeln("Existing: <<[model.Existing]>>");
            builder.Writeln("Missing: <<[model.Missing]>>"); // This member does not exist.

            // Prepare the data source.
            ReportModel model = new ReportModel { Existing = "Hello, World!" };

            // Configure the reporting engine to treat missing members as null.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Optional: customize the message shown for a plain missing member reference.
            engine.MissingMemberMessage = string.Empty;

            // Build the report. The missing member will be rendered as an empty string.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
