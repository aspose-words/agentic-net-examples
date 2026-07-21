using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesDemo
{
    // Simple data model with a single valid property.
    public class Model
    {
        public string ExistingProperty { get; set; } = "Valid value";
        // Note: No property named MissingProperty – this will trigger an inline error.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a valid tag.
            builder.Writeln("<<[model.ExistingProperty]>>");

            // Insert an invalid tag that references a non‑existent member.
            // With InlineErrorMessages enabled, the engine will embed an error message here.
            builder.Writeln("<<[model.MissingProperty]>>");

            // Prepare the data source.
            Model model = new Model();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The method returns false because there is a syntax error.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag.
            Console.WriteLine($"BuildReport success flag: {success}");

            // Save the resulting document for manual inspection (optional).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InlineErrorReport.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
