using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Extension method container. The method is static and will be called as a static method from the template.
    public static class DateTimeExtensions
    {
        // Formats the date according to the specified locale (culture name, e.g., "en-US").
        public static string FormatDate(DateTime date, string locale)
        {
            var culture = new CultureInfo(locale);
            // Long date pattern.
            return date.ToString("D", culture);
        }
    }

    // Simple data model for the report.
    public class Order
    {
        // Sample date property.
        public DateTime OrderDate { get; set; } = DateTime.Now;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Use static method syntax in the LINQ Reporting tags.
            builder.Writeln("Order date (US format): <<[DateTimeExtensions.FormatDate(order.OrderDate, \"en-US\")]>>");
            builder.Writeln("Order date (German format): <<[DateTimeExtensions.FormatDate(order.OrderDate, \"de-DE\")]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare the data source.
            var order = new Order
            {
                // Example specific date.
                OrderDate = new DateTime(2023, 12, 25)
            };

            // 4. Configure the ReportingEngine.
            var engine = new ReportingEngine();
            // Register the type that contains the static method.
            engine.KnownTypes.Add(typeof(DateTimeExtensions));

            // 5. Build the report.
            engine.BuildReport(doc, order, "order");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
