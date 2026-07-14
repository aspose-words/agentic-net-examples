using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingHtml
{
    // Model class that holds the HTML snippet to be inserted.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string HtmlSnippet { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the template document with a LINQ Reporting tag.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Report generated with an external HTML snippet:");
            // The <<html>> tag will be replaced by the value of HtmlSnippet at runtime.
            builder.Writeln("<<html [model.HtmlSnippet]>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Step 2: Load the template for reporting.
            var loadedTemplate = new Document(templatePath);

            // Step 3: Prepare the data model with the HTML content.
            var model = new ReportModel
            {
                HtmlSnippet = "<p style='color:blue;'>This is <b>HTML</b> inserted at runtime.</p>"
            };

            // Step 4: Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.
            engine.BuildReport(loadedTemplate, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "output.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
