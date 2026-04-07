using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model that provides the current UTC time.
    public class ReportModel
    {
        // Initialized with the current UTC time at the moment of object creation.
        public DateTime CurrentUtc { get; set; } = DateTime.UtcNow;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that will be replaced with the UTC time.
            // The root object name used later will be "model".
            builder.Writeln("Current UTC time: <<[model.CurrentUtc]>>");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
