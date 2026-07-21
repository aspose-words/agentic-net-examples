using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting engine.
    public class Invoice
    {
        // Price of the item.
        public decimal Price { get; set; } = 0m;

        // Tax rate (e.g., 0.20 for 20%).
        public decimal TaxRate { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write static text and LINQ Reporting tags.
            builder.Writeln("Price: <<[model.Price]>>");
            builder.Writeln("Tax Rate: <<[model.TaxRate]>>");
            // Calculate tax using an expression tag.
            builder.Writeln("Tax (Price * TaxRate): <<[model.Price * model.TaxRate]>>");

            // Save the template to a local file (required before BuildReport).
            const string templatePath = "InvoiceTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template back (simulating a real‑world scenario where the template exists on disk).
            Document loadedTemplate = new Document(templatePath);

            // 3. Prepare the data source.
            Invoice invoice = new Invoice
            {
                Price = 123.45m,
                TaxRate = 0.20m // 20%
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(loadedTemplate, invoice, "model");

            // 5. Save the generated report.
            const string outputPath = "InvoiceReport.docx";
            loadedTemplate.Save(outputPath);

            // Optional: inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
