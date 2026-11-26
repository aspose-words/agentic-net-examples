using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtmlExample
{
    // Simple data model containing an HTML snippet.
    public class ReportModel
    {
        // Sample HTML content to be inserted into the document.
        public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i> with a <span style=\"color:blue;\">blue span</span>.";
    }

    public class Program
    {
        public static void Main()
        {
            // Define file paths.
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputFolder);
            string templatePath = Path.Combine(outputFolder, "Template.docx");
            string resultPath = Path.Combine(outputFolder, "Result.docx");

            // -----------------------------------------------------------------
            // Step 1: Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("LINQ Reporting Example: HTML Insertion");
            builder.Writeln();

            // Insert the HTML expression tag using the -html switch.
            // The tag will be replaced with the evaluated HTML snippet.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            Document docToReport = new Document(templatePath);

            // Prepare the data model.
            ReportModel model = new ReportModel();

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(docToReport, model, "model");

            // Save the generated document.
            docToReport.Save(resultPath);
        }
    }
}
