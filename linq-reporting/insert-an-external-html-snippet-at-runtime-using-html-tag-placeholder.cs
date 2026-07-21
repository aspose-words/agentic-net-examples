using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // HTML snippet that will be inserted into the document at runtime.
        public string HtmlSnippet { get; set; } = "<p style='color:blue;'>Hello <b>World</b> from HTML snippet!</p>";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any required encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert the LINQ Reporting tag that will render the HTML snippet.
            // The tag uses the <<html>> syntax as required by the engine.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk (required before building the report).
            string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);

            // Optional: inform the user that the process completed.
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
