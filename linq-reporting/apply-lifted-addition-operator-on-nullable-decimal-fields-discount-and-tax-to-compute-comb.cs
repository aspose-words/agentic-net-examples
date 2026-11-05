using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLiftedAddition
{
    // Data model with nullable decimal fields.
    public class Order
    {
        // Sample discount value (nullable).
        public decimal? Discount { get; set; } = 5.5m;

        // Sample tax value (nullable).
        public decimal? Tax { get; set; } = 2.25m;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a heading.
            builder.Writeln("=== Order Report ===");
            builder.Writeln();

            // Insert a LINQ Reporting tag that uses the lifted addition operator
            // on the nullable fields Discount and Tax.
            // The expression will evaluate to null if either operand is null.
            builder.Writeln("Combined Discount + Tax: <<[order.Discount + order.Tax]>>");
            builder.Writeln();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Create sample data.
            Order order = new Order
            {
                Discount = 5.5m, // You can set to null to see null handling.
                Tax = 2.25m
            };

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the root object name "order".
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
