using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
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
            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a paragraph with a dynamic link tag.
            // The tag will be replaced with a hyperlink whose URL and display text
            // come from the data model (model.Url and model.LinkText).
            builder.Writeln("Visit our site: <<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            var model = new ReportModel
            {
                Url = "https://www.example.com",
                LinkText = "Example Site"
            };

            // Create the reporting engine and generate the report.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the final document.
            loadedTemplate.Save(reportPath);
        }
    }
}
