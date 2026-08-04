using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Simple data model representing an order.
    public class Order
    {
        // Initialize non‑nullable reference type to avoid warnings.
        public string CustomerName { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for legacy encodings (required by Aspose.Words in some scenarios).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            // The tag references a collection named 'orders' that will be passed to BuildReport.
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>, Amount: <<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source – a list of Order objects.
            // -----------------------------------------------------------------
            List<Order> orders = new()
            {
                new Order { CustomerName = "John Doe", Amount = 123.45m },
                new Order { CustomerName = "Jane Smith", Amount = 678.90m },
                new Order { CustomerName = "Bob Johnson", Amount = 250.00m }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the overload that specifies the collection name.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.

            // The third argument ('orders') is the name used inside the template tags.
            bool success = engine.BuildReport(reportDoc, orders, "orders");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Optional: indicate success (no interactive prompts required).
            Console.WriteLine(success
                ? $"Report generated successfully: {reportPath}"
                : "Report generation failed.");
        }
    }
}
