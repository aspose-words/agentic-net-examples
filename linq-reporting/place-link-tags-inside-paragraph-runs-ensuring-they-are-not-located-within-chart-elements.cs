using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;
using System.Text;

namespace LinqReportingLinkExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        public string Url { get; set; } = "";
        public string LinkText { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph containing a link tag (will be replaced by the engine).
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");

            // Insert a chart shape after the paragraph. The link tag must NOT be inside this chart.
            builder.InsertChart(ChartType.Column, 400, 300);

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data for the report.
            ReportModel model = new ReportModel
            {
                Url = "https://example.com",
                LinkText = "Visit Example"
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save(resultPath);
        }
    }
}
