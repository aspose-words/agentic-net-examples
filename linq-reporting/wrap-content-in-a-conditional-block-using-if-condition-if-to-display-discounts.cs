using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingIfConditionExample
{
    // Sample data model.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public decimal Discount { get; set; } = 15.5m;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with a conditional block.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("<<if [order.Discount > 0]>>Discount: $<<[order.Discount]>> <</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            Order order = new Order(); // Sample data.

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, order, "order");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
