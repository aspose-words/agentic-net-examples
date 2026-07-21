using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model with a non‑nullable property to avoid warnings.
    public class Product
    {
        public decimal Price { get; set; } = 0m;
    }

    // Custom external type whose static members can be used in the template.
    public static class MyHelper
    {
        // Formats a decimal value as currency.
        public static string FormatCurrency(decimal value) => $"${value:F2}";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Product price: <<[MyHelper.FormatCurrency(Price)]>>");
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare the data source.
            var product = new Product { Price = 123.45m };

            // 4. Configure the ReportingEngine.
            var engine = new ReportingEngine();
            // Register the custom external type so the template can call its static members without reflection.
            engine.KnownTypes.Add(typeof(MyHelper));

            // 5. Build the report. Use the overload without a data source name to reference members directly.
            engine.BuildReport(doc, product);

            // 6. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
