using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class that will hold the HTML string.
    public class ReportData
    {
        // The HTML to be inserted into the document.
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains the LINQ Reporting tag:
            //   <<[data.HtmlContent]:html(sourceStyles)>>
            // The "sourceStyles" switch tells the engine to keep the original HTML styles.
            Document template = new Document(@"C:\Templates\ReportTemplate.docx");

            // Prepare the data source with the HTML you want to inject.
            var data = new ReportData
            {
                HtmlContent = @"<p style='font-size:14pt;color:#2B579A;'>
                                   This is <b>dynamic</b> HTML inserted via LINQ Reporting.
                               </p>"
            };

            // Create the ReportingEngine and configure any required options.
            var engine = new ReportingEngine
            {
                // Example option – remove empty paragraphs after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The second overload allows the template to reference the
            // data source object itself using the name "data".
            engine.BuildReport(template, data, "data");

            // Save the resulting document.
            template.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
