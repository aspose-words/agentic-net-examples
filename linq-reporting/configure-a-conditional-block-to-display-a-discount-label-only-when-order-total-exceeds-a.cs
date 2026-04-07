using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model representing an order.
    public class Order
    {
        // Total amount of the order.
        public decimal Total { get; set; }

        // Constructor with a default value to avoid nullable warnings.
        public Order()
        {
            Total = 0m;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define file names.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write the order total placeholder.
            builder.Writeln("Order Total: <<[order.Total]>>");

            // Conditional block: show discount label only when Total > 100.
            builder.Writeln("<<if [order.Total > 100]>>Discount Applied<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Sample order where the total exceeds the threshold.
            Order order = new Order { Total = 150m };

            // -------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, order, "order");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
