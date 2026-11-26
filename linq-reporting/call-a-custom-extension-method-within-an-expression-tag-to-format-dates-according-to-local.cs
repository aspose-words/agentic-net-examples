using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDateFormatting
{
    // Data model with a DateTime property and a helper method for formatting.
    public class Order
    {
        public DateTime OrderDate { get; set; } = DateTime.Now;

        // Wrapper method that forwards to the extension method.
        public string GetFormattedDate(string locale) => OrderDate.FormatDate(locale);
    }

    // Extension method to format dates according to a locale string.
    public static class DateTimeExtensions
    {
        public static string FormatDate(this DateTime date, string locale)
        {
            if (string.IsNullOrWhiteSpace(locale))
                return date.ToString(CultureInfo.InvariantCulture);

            try
            {
                var culture = new CultureInfo(locale);
                return date.ToString(culture);
            }
            catch (CultureNotFoundException)
            {
                // Fallback to invariant culture if the locale is invalid.
                return date.ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data.
            var order = new Order
            {
                OrderDate = new DateTime(2023, 12, 25, 14, 30, 0)
            };

            // Create a template document programmatically.
            const string templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags that call the wrapper method.
            builder.Writeln("Order Date (en-US): <<[order.GetFormattedDate(\"en-US\")]>>");
            builder.Writeln("Order Date (fr-FR): <<[order.GetFormattedDate(\"fr-FR\")]>>");
            builder.Writeln("Order Date (invalid locale): <<[order.GetFormattedDate(\"invalid-locale\")]>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
