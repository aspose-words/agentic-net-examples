using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLinkExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL that the hyperlink will point to.
        public string Url { get; set; } = string.Empty;

        // Text displayed for the hyperlink.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a Word template programmatically and insert a link tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Dynamic hyperlink example:");
            // The <<link>> tag will be replaced with a hyperlink during report generation.
            // The tag references the root object named "model".
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------------------------------------------------------------
            // 2. Load the template back and build the report using LINQ Reporting.
            // ---------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data for the report.
            ReportModel model = new ReportModel
            {
                Url = "https://www.example.com",
                LinkText = "Visit Example.com"
            };

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used inside the template to reference the root object.
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report to the output file.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
