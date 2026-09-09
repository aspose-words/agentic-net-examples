using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Order
    {
        // Total amount of the order.
        public decimal Total { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "OrderTemplate.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OrderReport.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Static text.
            builder.Writeln("Order Summary");
            builder.Writeln("----------------");

            // Insert the total amount.
            builder.Writeln("Total: <<[order.Total]>>");

            // Conditional section – appears only when Total > 100.
            builder.Writeln("<<if [order.Total > 100]>>");
            builder.Writeln("Congratulations! This order qualifies for free shipping.");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data: an order with a total that exceeds the threshold.
            Order sampleOrder = new Order { Total = 150m };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name in the template is "order".
            engine.BuildReport(doc, sampleOrder, "order");

            // Save the generated report.
            doc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
