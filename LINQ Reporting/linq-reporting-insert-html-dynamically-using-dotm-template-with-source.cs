using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data source class that will be used by the LINQ Reporting engine.
    public class ReportData
    {
        // Property that holds the HTML fragment to be inserted into the document.
        public string HtmlContent { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOTM template that contains a reporting tag with the -sourceStyles switch,
            // e.g. <<[data.HtmlContent] -sourceStyles>>
            Document template = new Document(@"C:\Templates\ReportTemplate.dotm");

            // Prepare the data source. The HTML string can contain any valid HTML markup.
            var data = new ReportData
            {
                HtmlContent = @"
                    <p align='right'>Right aligned paragraph.</p>
                    <p><b>Bold text</b> and <i>italic text</i> inside a paragraph.</p>
                    <div align='center'>Centered <span style='color:#FF0000;'>red</span> text.</div>"
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: remove empty paragraphs that may be left after processing.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The second parameter is the data source object,
            // the third parameter is the name used inside the template (e.g. "data").
            engine.BuildReport(template, data, "data");

            // Save the generated document.
            template.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
