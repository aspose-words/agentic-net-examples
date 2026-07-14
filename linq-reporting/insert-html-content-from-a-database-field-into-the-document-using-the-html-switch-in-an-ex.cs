using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtml
{
    // Data model that mimics a database record containing HTML content.
    public class ReportModel
    {
        // Sample HTML stored in a database field.
        public string HtmlContent { get; set; } = "<h2 style='color:Blue;'>Hello from the database!</h2><p>This paragraph is <b>bold</b> and <i>italic</i>.</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document with a LINQ Reporting tag that
            //    inserts HTML using the -html switch.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("=== LINQ Reporting HTML Insertion Example ===");
            builder.Writeln();

            // Insert the HTML expression tag. The tag references the model's
            // HtmlContent property and uses the -html switch.
            builder.Writeln("<<[model.HtmlContent] -html>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a separate load step).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source (simulated database record).
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // The example finishes without waiting for user input.
        }
    }
}
