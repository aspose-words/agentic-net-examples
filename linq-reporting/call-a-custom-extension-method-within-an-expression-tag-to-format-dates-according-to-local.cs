using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Extension method used inside the LINQ Reporting expression tag.
    public static class DateExtensions
    {
        // Formats the date according to the specified locale (culture name, e.g., "en-US", "fr-FR").
        public static string ToLocaleString(this DateTime date, string locale)
        {
            var culture = new CultureInfo(locale);
            // Use short date pattern for the given culture.
            return date.ToString(culture.DateTimeFormat.ShortDatePattern, culture);
        }
    }

    // Simple data model that will be bound to the template.
    public class Order
    {
        public DateTime OrderDate { get; set; } = DateTime.Now;
        public string CustomerName { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                OrderDate = new DateTime(2023, 5, 1),
                CustomerName = "Alice Smith"
            };

            // Create a blank document that will serve as the template.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a paragraph with a LINQ Reporting expression tag that calls the custom extension method.
            // The tag will format the OrderDate according to the French locale.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order date (French format): <<[order.OrderDate.ToLocaleString(\"fr-FR\")]>>");
            builder.Writeln("Order date (US format): <<[order.OrderDate.ToLocaleString(\"en-US\")]>>");

            // Configure the reporting engine.
            var engine = new ReportingEngine
            {
                // Allow the engine to resolve the extension method.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(DateExtensions));

            // Build the report using the document template and the data source.
            // The root object name must match the name used in the template tags ("order").
            engine.BuildReport(doc, order, "order");

            // Ensure the output directory exists.
            var outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the generated report.
            var outputPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
