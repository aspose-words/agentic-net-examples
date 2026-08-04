using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalBlock
{
    // Data model used as the root object for the report.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public double Discount { get; set; }
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

            // Insert a simple line with a placeholder for the customer's name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Conditional block: show discount only when it is greater than zero.
            builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>>%<</if>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order
            {
                CustomerName = "John Doe",
                Discount = 15.0 // Change to 0 to see the block omitted.
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template tags is "order".
            engine.BuildReport(report, order, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
