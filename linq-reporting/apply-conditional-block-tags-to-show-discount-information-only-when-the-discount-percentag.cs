using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalBlock
{
    // Simple data model used by the template.
    public class Order
    {
        public int Id { get; set; } = 0;
        public double DiscountPercentage { get; set; } = 0.0;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            Order order = new Order
            {
                Id = 12345,
                DiscountPercentage = 15.0   // Change to 0 to see the block omitted.
            };

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = "template.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write static text and data tags.
            builder.Writeln("Order Report");
            builder.Writeln("Order ID: <<[order.Id]>>");

            // Conditional block: show discount only when it is greater than zero.
            builder.Writeln("<<if [order.DiscountPercentage > 0]>>Discount: <<[order.DiscountPercentage]>>%<</if>>");

            // Save the template so that it can be loaded later (required by the rule set).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport expects the root object name to match the tag prefix used in the template.
            engine.BuildReport(reportDoc, order, "order");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
