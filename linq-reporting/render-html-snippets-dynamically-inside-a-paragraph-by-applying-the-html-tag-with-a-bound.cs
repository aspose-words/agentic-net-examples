using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with an HTML snippet.
    public class ReportModel
    {
        // Initialize with sample HTML to avoid nullable warnings.
        public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the temporary template and final report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that will contain the HTML snippet.
            builder.Writeln("Paragraph start: ");
            // LINQ Reporting tag that injects HTML from the bound expression.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");
            builder.Writeln(" :Paragraph end.");

            // Save the template to disk as required by the workflow.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            // BuildReport overload with root name "model" to match the tag.
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
