using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model with nullable decimal fields and a combined property using lifted addition.
    public class Order
    {
        public decimal? Discount { get; set; } = null;
        public decimal? Tax { get; set; } = null;

        // The lifted addition operator returns null if either operand is null.
        public decimal? Combined => Discount + Tax;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string resultPath = "Result.docx";

            // -------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Combined value: <<[order.Combined]>>");
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // Prepare the data source.
            // -------------------------------------------------
            var order = new Order
            {
                Discount = 5.5m,
                Tax = 2.3m
            };

            // -------------------------------------------------
            // Build the report using Aspose.Words LINQ Reporting Engine.
            // -------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // -------------------------------------------------
            // Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(resultPath);

            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(resultPath)}");
        }
    }
}
