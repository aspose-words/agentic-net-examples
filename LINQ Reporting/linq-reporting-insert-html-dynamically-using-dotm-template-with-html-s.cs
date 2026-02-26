using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class that will be referenced from the DOTM template.
    // The template should contain a tag like <<[ds.Html]:html>> to insert the HTML.
    public class ReportData
    {
        public string Html { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains the LINQ Reporting tags.
            Document template = new Document(@"C:\Templates\ReportTemplate.dotm");

            // Prepare the data source with the HTML you want to insert dynamically.
            var data = new ReportData
            {
                Html = "<h2>Welcome</h2><p>This paragraph is <b>bold</b> and <i>italic</i>.</p>"
            };

            // Create the reporting engine and optionally set options.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may be left after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name ("ds") must match the name used in the template.
            engine.BuildReport(template, data, "ds");

            // Save the generated document.
            template.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
