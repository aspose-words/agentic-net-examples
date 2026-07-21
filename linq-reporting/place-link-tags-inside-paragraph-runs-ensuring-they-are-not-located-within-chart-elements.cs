using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL that the link will point to.
        public string Url { get; set; } = "https://www.example.com";

        // Text displayed for the hyperlink.
        public string LinkText { get; set; } = "Visit Example.com";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the temporary template and final report.
            const string templatePath = "Template.docx";
            const string outputPath = "ReportOutput.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title paragraph.
            builder.Writeln("LINQ Reporting – Link Example");

            // Insert a link tag inside a regular paragraph run.
            // The tag must not be placed inside any chart element.
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
