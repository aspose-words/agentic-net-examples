using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingHtmlExample
{
    // Simple data model with an HTML snippet.
    public class ReportModel
    {
        public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document with a LINQ Reporting tag.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a paragraph that will render the HTML snippet.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
