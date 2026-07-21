using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class used as the root data source for the report.
    public class ReportModel
    {
        // Sample date to be formatted.
        public DateTime ReportDate { get; set; } = DateTime.Now;
    }

    // Static class containing an extension method for DateTime formatting.
    // The method is registered with the ReportingEngine via KnownTypes.
    public static class DateExtensions
    {
        // Formats the date according to the specified locale (e.g., "en-US", "fr-FR").
        public static string FormatDate(this DateTime date, string locale)
        {
            var culture = new CultureInfo(locale);
            // Long date pattern for the given culture.
            return date.ToString("D", culture);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a simple template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // The expression calls the static method (registered in KnownTypes) to format the date.
            // Extension methods are not resolved directly on the instance type, so we invoke it as a static method.
            builder.Writeln("Report generated on: <<[DateExtensions.FormatDate(model.ReportDate, \"en-US\")]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Example date; you can set any specific date here.
                ReportDate = new DateTime(2023, 12, 25)
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(DateExtensions));

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
