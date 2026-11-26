using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Extension method for formatting decimal as currency string (kept for possible reuse).
    public static class CurrencyExtensions
    {
        public static string ToCurrencyString(this decimal amount) =>
            string.Format(CultureInfo.InvariantCulture, "${0:N2}", amount);
    }

    // Data model used as the root object for the report.
    public class Order
    {
        public decimal Amount { get; set; } = 0m;
        public string CustomerName { get; set; } = string.Empty;

        // Formatted amount string used by the LINQ Reporting template.
        public string AmountString => Amount.ToCurrencyString();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");

            // Create the template document with a LINQ Reporting tag.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Amount: <<[order.AmountString]>>");
            templateDoc.Save(templatePath);

            // Load the template (simulating a separate load step).
            var loadedTemplate = new Document(templatePath);

            // Create sample data.
            var order = new Order
            {
                CustomerName = "John Doe",
                Amount = 1234.56m
            };

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, order, "order");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}
