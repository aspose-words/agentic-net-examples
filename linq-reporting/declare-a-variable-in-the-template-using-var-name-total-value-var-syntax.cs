using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used as the root data source for the report.
    public class ReportModel
    {
        // Initialize the property to avoid nullable warnings.
        public int Total { get; set; } = 123;
    }

    public class Program
    {
        public static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create the template document programmatically.
            // ------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write a line that will display the value of the "Total" variable.
            // The expression <<[model.Total]>> references the "Total" property of the root object named "model".
            builder.Writeln("Total is: <<[model.Total]>>");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // ------------------------------------------------------------
            // 2. Load the template for reporting.
            // ------------------------------------------------------------
            Document report = new Document(templatePath);

            // ------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // ------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // No special options are required for this simple scenario.
                Options = ReportBuildOptions.None
            };

            // Create an instance of the data model.
            ReportModel model = new ReportModel();

            // Build the report. The root object name must match the name used in the template tags ("model").
            engine.BuildReport(report, model, "model");

            // ------------------------------------------------------------
            // 4. Save the generated report.
            // ------------------------------------------------------------
            const string reportPath = "Report.docx";
            report.Save(reportPath);
        }
    }
}
