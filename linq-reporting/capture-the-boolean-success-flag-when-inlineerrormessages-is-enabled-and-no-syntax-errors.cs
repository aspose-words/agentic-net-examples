using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a valid LINQ Reporting tag that references the model.
            builder.Writeln("Customer Name: <<[model.Name]>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the result flag.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
