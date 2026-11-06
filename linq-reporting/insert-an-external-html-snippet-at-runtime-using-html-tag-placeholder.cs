using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingHtmlInsert
{
    // Model class exposing the HTML snippet.
    public class ReportModel
    {
        public string HtmlSnippet { get; set; } = "<p style=\"color:blue;\">This is <strong>HTML</strong> inserted at runtime.</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document with an HTML placeholder.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a simple paragraph with the LINQ Reporting HTML tag.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report.
            // -------------------------------------------------
            var doc = new Document(templatePath);
            var model = new ReportModel();

            var engine = new ReportingEngine();
            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            doc.Save(outputPath);
        }
    }
}
