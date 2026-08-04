using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the final report
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically and insert the tag
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            // The <<html>> tag will be replaced with the HTML snippet from the model
            builder.Writeln("<<html [model.HtmlSnippet]>>");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template document for reporting
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // Prepare the data model containing the HTML snippet
            // -----------------------------------------------------------------
            var model = new ReportModel();

            // -----------------------------------------------------------------
            // Build the report using the LINQ Reporting engine
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // Save the generated report
            // -----------------------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }

    // Data model used by the template; property is initialized to avoid nullable warnings
    public class ReportModel
    {
        public string HtmlSnippet { get; set; } = "<p style='color:blue;'>This is <b>HTML</b> snippet inserted at runtime.</p>";
    }
}
