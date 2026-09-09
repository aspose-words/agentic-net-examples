using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model with a Total property.
    public class Order
    {
        public decimal Total { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a line showing the order total.
            builder.Writeln("Order total: <<[order.Total]>>");

            // Write a line showing the discount.
            // If Total > 500, display Total * 0.1m, otherwise display 0.
            // Note the use of the decimal literal (0.1m) to avoid type mismatch.
            builder.Writeln("Discount: " +
                "<<if [order.Total > 500]>>" +
                "<<[order.Total * 0.1m]>>" +
                "<</if>>" +
                "<<if [order.Total <= 500]>>0<</if>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order { Total = 620m }; // Example total > 500

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "order".
            engine.BuildReport(doc, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
