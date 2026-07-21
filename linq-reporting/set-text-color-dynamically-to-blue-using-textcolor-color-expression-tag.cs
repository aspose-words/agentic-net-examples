using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Color expression returned to the template. Any known color name or HTML code is accepted.
        public string Color { get; set; } = "Blue";

        // Text that will be displayed inside the colored region.
        public string Status { get; set; } = "Dynamic Status";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that sets the text color dynamically.
            // The color is taken from the model's Color property, and the inner text
            // is taken from the model's Status property.
            builder.Writeln("<<textColor [model.Color]>><<[model.Status]>> <</textColor>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel(); // Color = "Blue", Status = "Dynamic Status"

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            reportDoc.Save(reportPath);
        }
    }
}
