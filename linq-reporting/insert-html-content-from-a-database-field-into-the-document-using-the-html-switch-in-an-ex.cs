using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtml
{
    public class Program
    {
        public static void Main()
        {
            // Sample data model simulating a database record that contains HTML.
            ReportModel model = new ReportModel
            {
                HtmlContent = "<p style='color:blue;'>This is <b>HTML</b> content from DB.</p>"
            };

            // Create a template document and insert the LINQ Reporting HTML tag.
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<[model.HtmlContent] -html>>");
            templateDoc.Save(templatePath);

            // Load the template and build the report.
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document.
            reportDoc.Save("ReportWithHtml.docx");
        }
    }

    // Public data model with a property that holds HTML text.
    public class ReportModel
    {
        public string HtmlContent { get; set; } = string.Empty;
    }
}
