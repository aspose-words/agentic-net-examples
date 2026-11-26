using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model with price and tax rate.
    public class Invoice
    {
        public decimal Price { get; set; } = 123.45m;
        public decimal TaxRate { get; set; } = 0.08m; // 8%
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags with delimiters.
            builder.Writeln("Price: <<[Price]>>");
            builder.Writeln("Tax Rate: <<[TaxRate]>>");
            builder.Writeln("Tax Amount (Price * TaxRate): <<[Price * TaxRate]>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template (simulating a separate load step).
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            Invoice model = new Invoice();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
