using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL of the hyperlink.
        public string Url { get; set; } = string.Empty;

        // Text that will be displayed for the hyperlink.
        public string Text { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any legacy encodings Aspose.Words might need.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that creates a hyperlink.
            // The tag uses the data fields Url and Text from the model.
            builder.Writeln("<<link [model.Url] [model.Text]>>");

            // Save the template to disk (required before building the report).
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Url  = "https://www.example.com",
                Text = "Visit Example.com"
            };

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
