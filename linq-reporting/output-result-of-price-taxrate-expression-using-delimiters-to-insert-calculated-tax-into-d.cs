using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with price and tax rate.
    public class Invoice
    {
        // Initialize properties to avoid nullable warnings.
        public double Price { get; set; } = 199.99;
        public double TaxRate { get; set; } = 0.08; // 8 %
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags. The root object name will be "invoice".
            builder.Writeln("Price: <<[invoice.Price]>>");
            builder.Writeln("Tax Rate: <<[invoice.TaxRate]>>");
            // Calculate tax using an expression inside the delimiters.
            builder.Writeln("Tax Amount: <<[invoice.Price * invoice.TaxRate]>>");

            // Prepare the data source.
            Invoice invoice = new Invoice();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, invoice, "invoice");

            // Save the generated document.
            const string outputPath = "InvoiceReport.docx";
            doc.Save(outputPath);

            // Optionally inform the user where the file was saved.
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
