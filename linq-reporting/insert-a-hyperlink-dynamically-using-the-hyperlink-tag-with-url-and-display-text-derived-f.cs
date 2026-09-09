using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL for the hyperlink.
        public string Url { get; set; } = string.Empty;

        // Text that will be displayed as the hyperlink.
        public string Text { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting link tag.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // The <<link>> tag will be replaced with a hyperlink whose URL and
            // display text are taken from the data source fields Url and Text.
            builder.Writeln("<<link [Url] [Text]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                Url = "https://www.example.com",
                Text = "Visit Example"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // No special options are required for this simple scenario.
                Options = ReportBuildOptions.None
            };

            // The root object name in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
