using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

// Ensure code page support for Aspose.Words if needed.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

namespace LinqReportingForeachExample
{
    // Public data model classes.
    public class Invoice
    {
        public int Id { get; set; } = 0;
        public decimal Amount { get; set; } = 0m;
        public DateTime Date { get; set; } = DateTime.MinValue;
    }

    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Invoices = new List<Invoice>
                {
                    new Invoice { Id = 1, Amount = 123.45m, Date = new DateTime(2023, 1, 15) },
                    new Invoice { Id = 2, Amount = 678.90m, Date = new DateTime(2023, 2, 20) },
                    new Invoice { Id = 3, Amount = 250.00m, Date = new DateTime(2023, 3, 5) }
                }
            };

            // Create the template document programmatically.
            string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Invoices Report");
            builder.Writeln("<<foreach [Invoice invoice in Invoices]>>");
            builder.Writeln("Id: <<[invoice.Id]>>, Amount: <<[invoice.Amount]>>, Date: <<[invoice.Date]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
