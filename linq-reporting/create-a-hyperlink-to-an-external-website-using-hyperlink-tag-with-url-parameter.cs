using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL of the external website.
        public string Url { get; set; } = string.Empty;

        // Text displayed for the hyperlink.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with a <<link>> tag.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // The tag will be replaced with a hyperlink during report generation.
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");
            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template document.
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            ReportModel model = new ReportModel
            {
                Url = "https://www.example.com",
                LinkText = "Visit Example"
            };

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
