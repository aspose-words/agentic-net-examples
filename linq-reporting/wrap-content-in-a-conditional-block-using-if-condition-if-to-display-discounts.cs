using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the template.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public string CustomerName { get; set; } = string.Empty;
        public decimal Discount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple line with a placeholder for the customer's name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Conditional block: show the discount only when it is greater than zero.
            // The syntax must be exactly <<if [condition]>> ... <</if>>.
            builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>> <</if>>");

            // Save the template to disk so it can be loaded later.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data source.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Discount = 15.5m // Change to 0 to see the conditional block omitted.
            };

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
