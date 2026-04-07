using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with price and tax rate
    public class Invoice
    {
        public decimal Price { get; set; } = 0;
        public decimal TaxRate { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data
            var invoice = new Invoice
            {
                Price = 123.45m,
                TaxRate = 0.08m // 8%
            };

            // Create a template document programmatically
            const string templateFile = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Price: $<<[Price]>>");
            builder.Writeln("Tax Rate: <<[TaxRate]>>");
            builder.Writeln("Calculated Tax (Price * TaxRate): $<<[Price * TaxRate]>>");
            builder.Writeln("Total (Price + Tax): $<<[Price + (Price * TaxRate)]>>");

            // Save the template
            templateDoc.Save(templateFile);

            // Load the template and populate it using the LINQ Reporting Engine
            var reportDoc = new Document(templateFile);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, invoice); // No root name needed for direct member access

            // Save the final report
            const string reportFile = "Report.docx";
            reportDoc.Save(reportFile);
        }
    }
}
