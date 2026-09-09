using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionDemo
{
    // Extension methods must be defined in a static class.
    public static class DecimalExtensions
    {
        // Returns true if the amount exceeds the specified limit.
        public static bool IsHighValue(this decimal amount, decimal limit) => amount > limit;
    }

    // Data model used as the root object for the report.
    public class Order
    {
        // Sample amount property.
        public decimal Amount { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that calls the extension method on the Amount property.
            // The tag will output "True" or "False" depending on the limit (100 in this case).
            builder.Writeln("Amount: <<[order.Amount]>>");
            builder.Writeln("Is high (limit 100): <<[order.Amount.IsHighValue(100)]>>");

            // 2. Prepare the data source.
            Order order = new Order { Amount = 150m };

            // 3. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow the engine to use extension methods.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // 4. Build the report using the template, data source, and root name.
            engine.BuildReport(doc, order, "order");

            // 5. Save the generated report.
            doc.Save("Report_Output.docx");
        }
    }
}
