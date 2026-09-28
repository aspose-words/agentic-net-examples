using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingTimeSpanExample
{
    // Data model with a duration string and a parsed TimeSpan property.
    public class Order
    {
        // Sample duration string in the format "hh:mm:ss".
        public string DurationString { get; set; } = "02:30:45";

        // Parses the string into a TimeSpan using TimeSpan.Parse.
        public TimeSpan Duration => TimeSpan.Parse(DurationString);

        public string Description { get; set; } = "Sample order description";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a title.
            builder.Writeln("Order Report");
            builder.Writeln();

            // Insert LINQ Reporting tags that reference the model.
            builder.Writeln("Description: <<[order.Description]>>");
            builder.Writeln("Duration string: <<[order.DurationString]>>");
            builder.Writeln("Parsed TimeSpan: <<[order.Duration]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            Order order = new Order(); // Sample data.

            ReportingEngine engine = new ReportingEngine();
            // Build the report using the root object name "order".
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
