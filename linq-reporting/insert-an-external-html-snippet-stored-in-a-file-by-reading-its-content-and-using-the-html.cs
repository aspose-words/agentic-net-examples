using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtmlInsert
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Holds the HTML snippet that will be inserted into the document.
        public string HtmlSnippet { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare an external HTML file that will be read at runtime.
            // -----------------------------------------------------------------
            string htmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "snippet.html");
            const string sampleHtml = "<p style='color:blue; font-weight:bold;'>Hello from external HTML file!</p>";
            File.WriteAllText(htmlFilePath, sampleHtml);

            // -----------------------------------------------------------------
            // 2. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple text before the HTML insertion.
            builder.Writeln("Report generated with external HTML snippet:");
            // LINQ Reporting tag that inserts the HTML content using the -html switch.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            // Save the template to disk (required before building the report).
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Read the HTML snippet from the external file.
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Populate the model with the HTML content.
            ReportModel model = new ReportModel { HtmlSnippet = htmlContent };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);

            // Indicate successful completion (no interactive prompts).
            Console.WriteLine("Report generated successfully: " + outputPath);
        }
    }
}
