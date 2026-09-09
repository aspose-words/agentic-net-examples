using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Model class representing data that would normally come from a database.
    public class ReportModel
    {
        // HTML content stored in a database field.
        public string HtmlContent { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting tag that
            //    inserts HTML using the "-html" switch.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("=== LINQ Reporting HTML Insertion Example ===");
            // The tag below will be replaced with the HTML from the data source.
            builder.Writeln("<<[model.HtmlContent] -html>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document (simulating a separate load step).
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source. In a real scenario this would be read
            //    from a database; here we use a hard‑coded HTML snippet.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                HtmlContent = "<p style='color:blue; font-size:14pt;'>" +
                              "This is <b>HTML</b> content inserted from a data source." +
                              "</p>"
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);

            // Indicate successful completion (no interactive prompts).
            Console.WriteLine("Report generated successfully.");
        }
    }
}
