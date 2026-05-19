using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model containing a duration string and a parsed TimeSpan.
    public class Order
    {
        // Sample duration string in "hh:mm:ss" format.
        public string DurationString { get; set; } = "02:15:30";

        // Parses the string using TimeSpan.Parse.
        public TimeSpan Duration => TimeSpan.Parse(DurationString);
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that will display the parsed TimeSpan.
            builder.Writeln("Duration: <<[order.Duration]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            Order order = new Order(); // DurationString is already initialized.

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(loadedTemplate, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
