using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingHyperlinkExample
{
    // Data model for the report.
    public class ReportModel
    {
        // URL to link to.
        public string Url { get; set; } = string.Empty;

        // Text displayed for the hyperlink.
        public string LinkText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Paths for template and result documents.
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // -----------------------------------------------------------------
            // Create the template document with a LINQ Reporting link tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a simple paragraph containing the link tag.
            // The tag uses the model's Url and LinkText properties.
            builder.Writeln("Here is a dynamic hyperlink: <<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Sample data for the report.
            ReportModel model = new ReportModel
            {
                Url = "https://example.com",
                LinkText = "Visit Example"
            };

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated document.
            loadedTemplate.Save(resultPath);
        }
    }
}
