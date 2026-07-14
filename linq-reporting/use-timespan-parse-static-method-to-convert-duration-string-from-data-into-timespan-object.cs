using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Original duration string from the data source.
        public string DurationString { get; set; } = "01:30:45";

        // Converts the string to a TimeSpan using TimeSpan.Parse.
        public TimeSpan Duration => TimeSpan.Parse(DurationString);
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank Word document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple paragraph with a LINQ Reporting tag that references the TimeSpan property.
            builder.Writeln("Duration: <<[model.Duration]>>");

            // 2. Prepare the data source.
            ReportModel model = new ReportModel
            {
                // Example value; can be changed to test different inputs.
                DurationString = "02:15:30"
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model", so we pass it explicitly.
            engine.BuildReport(template, model, "model");

            // 4. Save the generated document.
            const string outputPath = "LinqReporting_TimeSpanReport.docx";
            template.Save(outputPath);
        }
    }
}
