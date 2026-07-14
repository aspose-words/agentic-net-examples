using System;
using System.Collections.Generic;
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

        // Initialize properties to avoid nullable warnings.
        public Invoice()
        {
            Id = 0;
            Amount = 0m;
            Date = DateTime.MinValue;
        }
    }

    // Wrapper class that holds the collection of invoices.
    public class ReportData
    {
        public List<Invoice> Invoices { get; set; } = new();

        // Sample data bootstrap.
        public ReportData()
        {
            Invoices.Add(new Invoice { Id = 1001, Amount = 250.75m, Date = new DateTime(2023, 1, 15) });
            Invoices.Add(new Invoice { Id = 1002, Amount = 480.00m, Date = new DateTime(2023, 2, 3) });
            Invoices.Add(new Invoice { Id = 1003, Amount = 125.50m, Date = new DateTime(2023, 3, 22) });
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a heading.
            builder.Writeln("Invoice Report");
            builder.Writeln();

            // Use the foreach tag (correct syntax) to iterate over the collection.
            builder.Writeln("<<foreach [inv in model.Invoices]>>");
            builder.Writeln("Invoice ID: <<[inv.Id]>>");
            builder.Writeln("Date: <<[inv.Date]>>");
            builder.Writeln("Amount: $<<[inv.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template document.
            var doc = new Document(templatePath);

            // 3. Prepare the data source.
            var data = new ReportData();

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, data, "model");

            // 5. Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
