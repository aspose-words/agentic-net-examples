using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingMinExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of prices.
        public List<decimal> Prices { get; set; } = new();

        // Lowest price calculated from the collection.
        public decimal LowestPrice => Prices.Min();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Prices = new List<decimal> { 199.99m, 149.50m, 179.75m, 129.99m, 159.00m }
            };

            // Create a template document programmatically.
            string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Discount benchmark price: <<[model.LowestPrice]>>");
            doc.Save(templatePath);

            // Load the template for reporting.
            var template = new Document(templatePath);

            // Build the report using LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            string outputPath = "report.docx";
            template.Save(outputPath);
        }
    }
}
