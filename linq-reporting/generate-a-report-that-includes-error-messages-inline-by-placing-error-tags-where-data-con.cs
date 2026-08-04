using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class OrderModel
    {
        // Customer name – valid string.
        public string CustomerName { get; set; } = "John Doe";

        // Order date – intentionally invalid to trigger a conversion error.
        public string OrderDate { get; set; } = "not-a-valid-date";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line that will be filled correctly.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Insert a line that attempts to format a date.
            // The provided value is not a valid date, so with InlineErrorMessages the engine will insert <<error>>.
            builder.Writeln("Order Date: <<[order.OrderDate]:date>>");

            // Prepare the data source.
            OrderModel order = new OrderModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The root object name must match the name used in the template tags ("order").
            bool success = engine.BuildReport(doc, order, "order");

            // Save the generated document.
            const string outputPath = "ReportWithInlineErrors.docx";
            doc.Save(outputPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report built successfully: {success}");
            Console.WriteLine($"Output saved to: {outputPath}");
        }
    }
}
