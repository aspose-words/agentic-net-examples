using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the LINQ Reporting template.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public double Total { get; set; } = 150.0;
        public double Discount { get; set; } = 10.0; // Percentage; set to 0 to hide the block.
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order total: $<<[order.Total]>>");
            // Conditional block: show only when Discount > 0.
            builder.Writeln("<<if [order.Discount > 0]>>Discount applied: <<[order.Discount]>>%<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required by the workflow).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order(); // Discount is 10%, so the block will appear.

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple scenario.
            engine.BuildReport(doc, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
