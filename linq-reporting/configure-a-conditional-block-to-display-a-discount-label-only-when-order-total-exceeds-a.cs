using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingConditionalDiscount
{
    // Public data model for the order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public decimal Total { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "order_template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "order_report.docx");

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Order Summary");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.Total]>>");
            // Conditional block: show discount label only when Total > 100.
            builder.Writeln("<<if [order.Total > 100]>>Discount Applied<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Total = 150m // Exceeds the threshold, so the discount label will appear.
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
