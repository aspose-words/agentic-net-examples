using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model representing an invoice.
    public class Invoice
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
    }

    // Wrapper class that holds the collection used by the report.
    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Invoices.Add(new Invoice { Id = 1, Amount = 123.45m });
            model.Invoices.Add(new Invoice { Id = 2, Amount = 678.90m });
            model.Invoices.Add(new Invoice { Id = 3, Amount = 250.00m });

            // -----------------------------------------------------------------
            // Create the template document containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templateFile = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Invoice Report");
            // Correct foreach syntax: variable name without explicit type.
            builder.Writeln("<<foreach [invoice in Invoices]>>");
            builder.Writeln("Id: <<[invoice.Id]>>   Amount: <<[invoice.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // Load the template and generate the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templateFile);
            var engine = new ReportingEngine();

            // Build the report using the model as the data source.
            engine.BuildReport(reportDoc, model);

            // Save the final report.
            reportDoc.Save("InvoiceReport.docx");
        }
    }
}
