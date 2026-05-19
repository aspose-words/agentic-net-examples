using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtmlInsert
{
    // Model class used as the root data source for the report.
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
            // 1. Prepare the external HTML snippet file.
            // -----------------------------------------------------------------
            const string htmlFileName = "snippet.html";
            const string htmlContent = "<p style='color:blue;'>Hello <b>World</b> from HTML snippet.</p>";
            File.WriteAllText(htmlFileName, htmlContent);

            // -----------------------------------------------------------------
            // 2. Create the template document programmatically.
            //    The template contains a LINQ Reporting tag that references the model's HtmlSnippet property
            //    and uses the -html switch to render the string as HTML.
            // -----------------------------------------------------------------
            const string templateFileName = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert the LINQ Reporting tag.
            builder.Writeln("<<[model.HtmlSnippet] -html>>");

            // Save the template to disk.
            templateDoc.Save(templateFileName);

            // -----------------------------------------------------------------
            // 3. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templateFileName);

            // -----------------------------------------------------------------
            // 4. Load the HTML snippet from the external file and populate the model.
            // -----------------------------------------------------------------
            string htmlFromFile = File.ReadAllText(htmlFileName);
            ReportModel model = new ReportModel { HtmlSnippet = htmlFromFile };

            // -----------------------------------------------------------------
            // 5. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the final document.
            // -----------------------------------------------------------------
            const string outputFileName = "Report.docx";
            reportDoc.Save(outputFileName);

            // Optional: indicate completion.
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputFileName)}");
        }
    }
}
