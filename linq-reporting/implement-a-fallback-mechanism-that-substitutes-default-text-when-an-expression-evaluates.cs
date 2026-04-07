using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingFallback
{
    // Simple data model used by the template.
    public class Order
    {
        // This property may be null; the engine will replace it with a default text.
        public string? CustomerName { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that references a possibly‑null property.
            // The tag will be replaced with the value of Order.CustomerName,
            // or with the fallback text defined in MissingMemberMessage.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Prepare the data source with a null value.
            Order order = new Order
            {
                CustomerName = null // Intentionally null to trigger the fallback.
            };

            // Configure the reporting engine to treat missing (null) members as null literals.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Text that will be inserted when the expression evaluates to null.
            engine.MissingMemberMessage = "N/A";

            // Build the report. The root object name must match the name used in the template tags.
            engine.BuildReport(doc, order, "order");

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
