using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model class that holds a duration string and parses it to a TimeSpan.
    public class Order
    {
        // Sample duration string in the format hh:mm:ss.
        public string DurationString { get; set; } = "01:30:45";

        // Parses the string into a TimeSpan using TimeSpan.Parse.
        public TimeSpan Duration => TimeSpan.Parse(DurationString);
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that will display the parsed TimeSpan.
            // The tag references the 'Duration' property of the root object named 'order'.
            builder.Writeln("Parsed duration: <<[order.Duration]>>");

            // Save the template to a local file (optional, but follows lifecycle rules).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Prepare the data source.
            Order order = new Order(); // DurationString = "01:30:45", Duration = 01:30:45

            // 3. Build the report using Aspose.Words ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Load the template document (demonstrating load step).
            Document doc = new Document(templatePath);
            // Build the report; the root object name must match the tag reference.
            engine.BuildReport(doc, order, "order");

            // 4. Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
