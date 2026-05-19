using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        // Note: No property named "Missing" – this will trigger an inline error.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Customer Report");
            builder.Writeln();

            // Valid tag – will be replaced with the Name value.
            builder.Writeln("Name: <<[model.Name]>>");

            // Invalid tag – property does not exist, will be shown as an inline error message.
            builder.Writeln("Missing property: <<[model.Missing]>>");

            // Conditional tag – will be empty because the condition is false (Age = 25).
            builder.Writeln("<<if [model.Age > 30]>>");
            builder.Writeln("Age over 30: <<[model.Age]>>");
            builder.Writeln("<</if>>");

            // Build the data source.
            ReportModel model = new ReportModel
            {
                Name = "John Doe",
                Age = 25
            };

            // Configure the reporting engine with both options.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned flag indicates whether parsing succeeded (relevant for InlineErrorMessages).
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag to the console (optional, does not require user interaction).
            Console.WriteLine($"Report build success: {success}");

            // Save the generated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
