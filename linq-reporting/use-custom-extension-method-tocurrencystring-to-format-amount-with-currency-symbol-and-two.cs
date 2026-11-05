using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Data model
    public class Order
    {
        public decimal Amount { get; set; } = 0m;

        // Exposes the formatted amount using the extension method.
        public string AmountFormatted => Amount.ToCurrencyString();
    }

    // Extension method for formatting currency.
    public static class DecimalExtensions
    {
        public static string ToCurrencyString(this decimal amount) => $"${amount:F2}";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Order Report");
            // Use the formatted property in the LINQ Reporting tag.
            builder.Writeln("Amount: <<[order.AmountFormatted]>>");

            // Save the template and reload it as required by the engine.
            const string templatePath = "Template.docx";
            template.Save(templatePath);
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Prepare sample data.
            // -----------------------------------------------------------------
            var order = new Order { Amount = 1234.56m };

            // -----------------------------------------------------------------
            // Build the report.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
