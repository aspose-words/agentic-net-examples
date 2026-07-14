using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExtensionExample
{
    // Data model used as the root object for the report.
    public class Order
    {
        public decimal Amount { get; set; }

        public Order(decimal amount) => Amount = amount;
    }

    // Static class that contains the custom extension method.
    public static class DecimalExtensions
    {
        // Returns true if the amount exceeds the specified limit.
        public static bool IsHighValue(this decimal amount, decimal limit) => amount > limit;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert tags that will be processed by the LINQ Reporting engine.
            // <<[order.Amount]>>      – displays the raw amount.
            // <<[order.Amount.IsHighValue(100)]>> – calls the custom extension method.
            builder.Writeln("Amount: <<[order.Amount]>>");
            builder.Writeln("Is high (>100): <<[order.Amount.IsHighValue(100)]>>");

            // Save the template to disk (required before loading it for reporting).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order(150m); // Sample order with an amount of 150.

            // -----------------------------------------------------------------
            // 4. Configure and run the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the static class that contains the extension method so that
            // the engine can resolve it during expression evaluation.
            engine.KnownTypes.Add(typeof(DecimalExtensions));

            // Allow the engine to resolve extension methods on value types.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Build the report. The root object name must match the name used in the tags.
            engine.BuildReport(reportDoc, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
