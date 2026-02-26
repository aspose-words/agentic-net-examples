using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source containing the HTML string to be inserted.
    public class ReportData
    {
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the PDF template that contains a LINQ Reporting tag, e.g. <<[data.HtmlContent]:html>>
            Document template = new Document("Template.pdf");

            // Prepare the data source with the HTML you want to inject.
            var data = new ReportData
            {
                HtmlContent = "<h1>Dynamic Title</h1><p>This paragraph is inserted <b>as HTML</b> at runtime.</p>"
            };

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used in the template to reference the data source.
            engine.BuildReport(template, data, "data");

            // Save the populated document as PDF.
            template.Save("Result.pdf");
        }
    }
}
