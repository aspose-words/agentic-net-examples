using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingForeachExample
{
    // Simple data model representing an invoice.
    public class Invoice
    {
        public int Id { get; set; }
        public decimal Amount { get; set; }
        public string Customer { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("Invoice Report");
            builder.Writeln();

            // Insert a foreach tag that iterates over the Invoices collection.
            // The tag follows the exact syntax required by Aspose.Words LINQ Reporting Engine.
            builder.Writeln("<<foreach [invoice in Invoices]>>");

            // Inside the loop write invoice fields.
            builder.Writeln("Id: <<[invoice.Id]>>");
            builder.Writeln("Customer: <<[invoice.Customer]>>");
            builder.Writeln("Amount: $<<[invoice.Amount]>>");
            builder.Writeln(); // Blank line between records.

            // Close the foreach block.
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Invoices = new List<Invoice>
                {
                    new Invoice { Id = 101, Customer = "Acme Corp", Amount = 1234.56m },
                    new Invoice { Id = 102, Customer = "Globex Inc", Amount = 7890.12m },
                    new Invoice { Id = 103, Customer = "Soylent Co", Amount = 345.67m }
                }
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated document.
            doc.Save("InvoiceReport.docx");
        }
    }
}
