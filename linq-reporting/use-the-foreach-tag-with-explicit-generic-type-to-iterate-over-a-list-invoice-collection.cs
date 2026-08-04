using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model representing an invoice.
    public class Invoice
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    // Wrapper class that holds the collection used in the template.
    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words for some encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var invoices = new List<Invoice>
            {
                new() { Id = 1001, Amount = 250.75m, Date = new DateTime(2023, 5, 12) },
                new() { Id = 1002, Amount = 480.00m, Date = new DateTime(2023, 6, 3) },
                new() { Id = 1003, Amount = 125.50m, Date = new DateTime(2023, 7, 21) }
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Invoice Report");
            builder.Writeln("----------------");

            // Correct foreach tag syntax (type is not specified in LINQ Reporting tags).
            builder.Writeln("<<foreach [inv in Invoices]>>");
            builder.Writeln("ID: <<[inv.Id]>>, Amount: <<[inv.Amount]>>, Date: <<[inv.Date]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // Load the template for report generation.
            var templateDoc = new Document(templatePath);

            // Prepare the root data object.
            var model = new ReportModel { Invoices = invoices };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            templateDoc.Save(outputPath);
        }
    }
}
