using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalDiscount
{
    // Data model for the report.
    public class Order
    {
        // Sample properties – initialized to avoid nullable warnings.
        public string CustomerName { get; set; } = "John Doe";
        public decimal Total { get; set; } = 0m;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write some static text and insert LINQ Reporting tags.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order Total: <<[order.Total]>>");

            // Conditional block – the discount label appears only when Total > 100.
            builder.Writeln("<<if [order.Total > 100]>>Discount Applied: 10%<</if>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order
            {
                CustomerName = "Alice Smith",
                Total = 150m // Change this value to test the conditional block.
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport expects the root object name to match the tag prefix ("order").
            engine.BuildReport(report, order, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            report.Save(reportPath);
        }
    }
}
