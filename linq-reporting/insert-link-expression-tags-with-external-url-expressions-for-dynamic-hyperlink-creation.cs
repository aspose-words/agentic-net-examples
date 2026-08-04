using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingLinkExample
{
    // Simple data model containing a URL and the text to display.
    public class ReportModel
    {
        public string Url { get; set; } = "https://www.example.com";
        public string LinkText { get; set; } = "Example Site";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for legacy encodings (required by Aspose.Words in some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a paragraph containing the LINQ Reporting link tag.
            // The tag will be replaced with a hyperlink whose URL and display text
            // come from the data model.
            builder.Writeln("Visit: <<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Create the data source.
            ReportModel model = new ReportModel();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final document.
            doc.Save(reportPath);
        }
    }
}
