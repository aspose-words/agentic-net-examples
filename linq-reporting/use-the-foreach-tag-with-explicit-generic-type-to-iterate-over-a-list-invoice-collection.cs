using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingForeachExample
{
    // Data model representing an invoice.
    public class Invoice
    {
        public int Id { get; set; }
        public string Customer { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    // Wrapper class required by the ReportingEngine to expose the collection.
    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "InvoiceTemplate.docx";
            const string reportPath = "InvoiceReport.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Invoice Report");
            // Correct foreach syntax: no explicit type inside the tag.
            builder.Writeln("<<foreach [inv in Invoices]>>");
            builder.Writeln("Id: <<[inv.Id]>>, Customer: <<[inv.Customer]>>, Amount: <<[inv.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Invoices = new List<Invoice>
                {
                    new Invoice { Id = 1, Customer = "Acme Corp", Amount = 1234.56m },
                    new Invoice { Id = 2, Customer = "Globex Inc", Amount = 7890.12m },
                    new Invoice { Id = 3, Customer = "Soylent Co", Amount = 345.67m }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name ("model") must match the name used in the template tags.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
