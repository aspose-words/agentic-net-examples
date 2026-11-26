using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingLiftedAddition
{
    // Data model with nullable decimal fields.
    public class Order
    {
        public decimal? Discount { get; set; } = 0m;
        public decimal? Tax { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var order = new Order
            {
                Discount = 5.5m,
                Tax = 2.3m
            };

            // Create a template document programmatically.
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Discount: <<[order.Discount]>>");
            builder.Writeln("Tax: <<[order.Tax]>>");
            // Use lifted addition operator on nullable decimals.
            builder.Writeln("Combined (Discount + Tax): <<[order.Discount + order.Tax]>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for reporting.
            var templateDoc = new Document(templatePath);

            // Build the report using LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, order, "order");

            // Save the generated report.
            var outputPath = "report.docx";
            templateDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
