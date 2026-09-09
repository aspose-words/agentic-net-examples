using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingFallback
{
    // Simple data model with a nullable property.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public string? CustomerName { get; set; } = null;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that references the CustomerName property.
            // If the property is null, we want to display a default text.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Prepare the data source with a null value.
            Order order = new Order
            {
                CustomerName = null // Simulate missing data.
            };

            // Configure the reporting engine to treat missing members as null
            // and replace them with a custom message.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            engine.MissingMemberMessage = "N/A";

            // Build the report. The root object name must match the tag prefix.
            engine.BuildReport(doc, order, "order");

            // Save the result to the working directory.
            doc.Save("ReportWithFallback.docx");
        }
    }
}
