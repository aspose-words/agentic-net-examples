using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Dynamic Heading";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure code pages are available (required for some data sources).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // ---------- Create the template document ----------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Set the heading font size to 14 points.
            builder.Font.Size = 14;

            // Insert a LINQ Reporting tag that will be replaced by the model's Title.
            builder.Writeln("<<[model.Title]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------- Prepare the data ----------
            ReportModel model = new ReportModel
            {
                Title = "Report Heading – Fourteen Point Font"
            };

            // ---------- Build the report ----------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data.
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            reportDoc.Save(outputPath);
        }
    }
}
