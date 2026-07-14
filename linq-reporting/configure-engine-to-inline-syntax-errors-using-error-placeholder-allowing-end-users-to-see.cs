using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingInlineErrors
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a valid tag that will be replaced with the person's name.
            builder.Writeln("Customer: <<[model.Name]>>");

            // Insert an invalid tag (property does not exist) to trigger an inline error message.
            builder.Writeln("Invalid field: <<[model.Unknown]>>");

            // Prepare the data source.
            Person model = new Person { Name = "John Doe" };

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The returned flag indicates whether parsing succeeded.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag to the console.
            Console.WriteLine($"Report build success: {success}");

            // Save the resulting document.
            doc.Save("ReportWithInlineErrors.docx");
        }
    }
}
