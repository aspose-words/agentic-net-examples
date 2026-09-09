using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalDiscount
{
    // Data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public double Total { get; set; } = 250.0;
        public double DiscountPercentage { get; set; } = 15.0; // Set to 0 to hide discount block.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "DiscountTemplate.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "DiscountReport.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple report layout with a conditional block that shows discount only when > 0.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.Total]>>");
            builder.Writeln("<<if [order.DiscountPercentage > 0]>>Discount: <<[order.DiscountPercentage]>>%<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required by the workflow).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Order sampleOrder = new Order(); // Uses the default values defined above.

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this example.
            engine.BuildReport(loadedTemplate, sampleOrder, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // The example finishes without waiting for user input.
        }
    }
}
