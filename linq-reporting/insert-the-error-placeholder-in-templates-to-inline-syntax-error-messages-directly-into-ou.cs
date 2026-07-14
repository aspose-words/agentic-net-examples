using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesExample
{
    // Simple data model with a single property.
    public class Customer
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare the data source.
            var model = new Customer();

            // Create a new blank document and insert LINQ Reporting tags.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Correct tag – will be replaced with the customer's name.
            builder.Writeln("Customer: <<[model.Name]>>");

            // Incorrect tag – references a non‑existent member and will generate an error.
            builder.Writeln("Missing property: <<[model.Missing]>>");

            // Inline error placeholder – will be replaced with the error message produced above.
            builder.Writeln("Error details: <<error>>");

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The third argument is the name used in the template to reference the root object.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save("ReportWithInlineErrors.docx");

            // Output the success flag (optional, not required for the example to run).
            Console.WriteLine($"Report generation successful: {success}");
        }
    }
}
