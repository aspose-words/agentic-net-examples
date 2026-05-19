using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the LINQ Reporting template.
    public class Order
    {
        // Non‑nullable string initialized to avoid CS8618.
        public string CustomerName { get; set; } = string.Empty;
        public double Discount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Simple static text.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Conditional block: show discount only when it is greater than zero.
            builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>>%<</if>>");

            // Save the template to disk (required by the workflow).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Sample data where Discount > 0 (the block will be rendered).
            Order order = new Order
            {
                CustomerName = "John Doe",
                Discount = 15.0
            };

            // Build the report using the root object name "order".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, order, "order");

            // Save the generated report.
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
