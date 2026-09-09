using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Text that will be displayed in the document.
        public string Status { get; set; } = "Pending";

        // Color expression for the textColor tag.
        // Can be a known color name, HTML hex code, or any value accepted by Aspose.Words.
        public string StatusColor { get; set; } = "Red";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains the textColor tag.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The tag applies a dynamic color to the run that follows it.
            // <<textColor [StatusColor]>> starts the colored region,
            // <<[Status]>> inserts the status text,
            // <</textColor>> ends the colored region.
            builder.Writeln("<<textColor [StatusColor]>><<[Status]>><</textColor>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Status = "Completed",
                StatusColor = "Green"
            };

            // Use the LINQ Reporting engine to populate the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model); // No data source name needed; members are accessed directly.

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
