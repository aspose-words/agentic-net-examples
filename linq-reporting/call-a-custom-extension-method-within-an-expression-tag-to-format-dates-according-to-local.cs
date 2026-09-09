using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Extension method used in the template expression.
    public static class DateExtensions
    {
        // Formats the given DateTime according to the specified locale (culture name).
        public static string Format(this DateTime date, string locale)
        {
            var culture = new CultureInfo(locale);
            // Example format: full date pattern of the culture.
            return date.ToString(culture.DateTimeFormat.LongDatePattern, culture);
        }
    }

    // Sample data model.
    public class Order
    {
        public DateTime OrderDate { get; set; } = DateTime.Now;
    }

    // Wrapper root object for the report.
    public class ReportModel
    {
        public Order Order { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a line that uses the custom extension method to format the date.
            // The expression calls DateExtensions.Format(date, locale).
            builder.Writeln("Order date (French locale): <<[DateExtensions.Format(Order.OrderDate, \"fr-FR\")]>>");
            builder.Writeln("Order date (Japanese locale): <<[DateExtensions.Format(Order.OrderDate, \"ja-JP\")]>>");

            // Save the template to disk (required before building the report).
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -------------------------------------------------
            var doc = new Document(templatePath);
            var model = new ReportModel(); // Root object with sample data.

            // -------------------------------------------------
            // 3. Configure and run the LINQ Reporting engine.
            // -------------------------------------------------
            var engine = new ReportingEngine();

            // Register the static class that contains the extension method so the engine can invoke it.
            engine.KnownTypes.Add(typeof(DateExtensions));

            // Build the report using the root object name "model".
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
