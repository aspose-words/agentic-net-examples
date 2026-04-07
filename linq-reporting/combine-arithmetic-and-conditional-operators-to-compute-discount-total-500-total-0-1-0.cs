using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        // Total amount of the order.
        public double Total { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a line that shows the order total.
            builder.Writeln("Total: <<[order.Total]>>");

            // Write a line that computes the discount using a conditional expression:
            // If Total > 500 then 10% of Total, otherwise 0.
            builder.Writeln("Discount: <<[order.Total > 500 ? order.Total * 0.1 : 0]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            Order sampleOrder = new Order { Total = 620.0 };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name used in the template is "order".
            engine.BuildReport(reportDoc, sampleOrder, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
