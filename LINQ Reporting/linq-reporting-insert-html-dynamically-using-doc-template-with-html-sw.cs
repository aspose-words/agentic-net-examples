using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data class that will be used as the data source for the report.
    public class ReportData
    {
        // The HTML string that will be inserted into the document.
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains the HTML switch.
            // The template should have a tag like <<[ds.HtmlContent]:html>> where the HTML will be inserted.
            Document template = new Document("TemplateWithHtmlSwitch.docx");

            // Prepare the data source with the HTML you want to inject.
            var data = new ReportData
            {
                HtmlContent = "<h2 style='color:blue;'>Dynamic Title</h2>" +
                              "<p>This paragraph is <b>bold</b> and <i>italic</i>.</p>" +
                              "<ul><li>Item 1</li><li>Item 2</li></ul>"
            };

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("ds") must match the name used in the template tag.
            engine.BuildReport(template, data, "ds");

            // Save the populated document.
            template.Save("ReportWithDynamicHtml.docx");
        }
    }
}
