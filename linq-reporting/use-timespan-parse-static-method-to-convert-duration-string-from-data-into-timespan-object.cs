using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Duration expressed as a string, e.g. "02:30:45".
        public string DurationString { get; set; } = "00:00:00";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that parses the duration string using TimeSpan.Parse.
            // The ReportingEngine must know the TimeSpan type to allow static method calls.
            builder.Writeln("Parsed duration: <<[TimeSpan.Parse(model.DurationString)]>>");

            // Save the template to disk.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                DurationString = "02:30:45" // 2 hours, 30 minutes, 45 seconds.
            };

            // -----------------------------------------------------------------
            // 3. Build the report.
            // -----------------------------------------------------------------
            // Load the template document.
            Document reportDoc = new Document(templatePath);

            // Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Register TimeSpan so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(TimeSpan));

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
