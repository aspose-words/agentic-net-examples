using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtmlExample
{
    // Simple data model containing an HTML snippet.
    public class ReportModel
    {
        // Initialize with a formatted HTML fragment.
        public string HtmlSnippet { get; set; } = "<p style='color:red; font-size:14pt;'>Hello <b>World</b> from <i>HTML</i>!</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert the LINQ Reporting tag that will embed the HTML snippet.
            // The "-html" switch tells the engine to treat the expression result as HTML.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document.
            reportDoc.Save(outputPath);
        }
    }
}
