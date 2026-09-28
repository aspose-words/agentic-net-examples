using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;
using System.Text;

namespace LinkTagExample
{
    // Data model for the report.
    public class ReportModel
    {
        public string Url { get; set; } = "";
        public string LinkText { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph with a link tag (will be replaced by the engine).
            builder.Writeln("Visit the site: <<link [model.Url] [model.LinkText]>>");

            // Insert a chart after the paragraph (link tags must NOT be inside charts).
            builder.InsertChart(ChartType.Column, 400, 300);

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel
            {
                Url = "https://example.com",
                LinkText = "Example Site"
            };

            // Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
