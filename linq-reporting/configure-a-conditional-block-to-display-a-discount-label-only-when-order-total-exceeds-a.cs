using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model for the report.
    public class Order
    {
        // Customer name – initialized to avoid nullable warnings.
        public string CustomerName { get; set; } = string.Empty;

        // Order total – using decimal for monetary values.
        public decimal Total { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document and insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Plain fields.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.Total]>>");

            // Conditional block – the label appears only when Total > 100.
            // Note the correct closing tag syntax: <</if>>.
            builder.Writeln("<<if [order.Total > 100]>>Discount Applied! Total = <<[order.Total]>> <</if>>");

            // 2. Prepare sample data.
            Order sampleOrder = new Order
            {
                CustomerName = "John Doe",
                Total = 150m // Change this value to test the condition.
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };

            // The root object name in the template is "order", so we pass the name accordingly.
            engine.BuildReport(doc, sampleOrder, "order");

            // 4. Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
