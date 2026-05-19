using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for the report.
    public class Invoice
    {
        // Initialize properties to avoid nullable warnings.
        public decimal Price { get; set; } = 0m;
        public decimal TaxRate { get; set; } = 0m;
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
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Invoice");
            builder.Writeln("------------------------------");
            builder.Writeln("Price: <<[invoice.Price]>>");
            builder.Writeln("Tax Rate: <<[invoice.TaxRate]>>");
            // Calculate tax using an expression tag.
            builder.Writeln("Tax Amount (Price * TaxRate): <<[invoice.Price * invoice.TaxRate]>>");
            builder.Writeln("------------------------------");
            builder.Writeln("Total: <<[invoice.Price + (invoice.Price * invoice.TaxRate)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Sample data.
            var invoice = new Invoice
            {
                Price = 199.99m,
                TaxRate = 0.07m // 7% tax
            };

            // Create the reporting engine.
            var engine = new ReportingEngine();

            // Build the report using the root object name "invoice".
            engine.BuildReport(reportDoc, invoice, "invoice");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
