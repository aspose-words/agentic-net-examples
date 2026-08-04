using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalExample
{
    // Simple data model representing an order.
    public class Order
    {
        // Customer name – initialized to avoid nullable warnings.
        public string CustomerName { get; set; } = "John Doe";

        // Order total – non‑nullable decimal.
        public decimal Total { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert placeholders for data fields.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.Total]>>");

            // Conditional block – the label appears only when Total > 100.
            builder.Writeln("<<if [order.Total > 100]>>Discount Applied<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            Order sampleOrder = new Order
            {
                CustomerName = "Alice Smith",
                Total = 150.75m // Change this value to test the condition.
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "order".
            engine.BuildReport(reportDoc, sampleOrder, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
