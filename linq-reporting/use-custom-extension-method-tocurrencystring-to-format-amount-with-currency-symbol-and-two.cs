using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Extension method for formatting currency.
    public static class Extensions
    {
        public static string ToCurrencyString(this decimal amount) => $"${amount:F2}";
    }

    // Simple data model.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public decimal Amount { get; set; } = 1234.56m;

        // Uses the extension method to provide a formatted string for the report.
        public string FormattedAmount => Amount.ToCurrencyString();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words on some platforms).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var order = new Order();

            // Create a Word template programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.FormattedAmount]>>");
            builder.Document.Save(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
