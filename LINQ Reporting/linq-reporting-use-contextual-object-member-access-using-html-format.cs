using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingHtmlExample
{
    // Data source class that contains an HTML string.
    public class ReportData
    {
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting Engine tag that will render the HTML content.
            // The ":html" format tells the engine to treat the string as HTML.
            builder.Writeln("<<[report.HtmlContent]:html>>");

            // Prepare the data source with some HTML markup.
            var data = new ReportData
            {
                HtmlContent = "<b>Hello</b> <i>World</i>!<br/><span style='color:blue;'>Blue text</span>"
            };

            // Build the report. The second parameter is the data source object,
            // the third parameter is the name used to reference it inside the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "report");

            // Save the resulting document.
            doc.Save("LinqReportingHtmlResult.docx");
        }
    }
}
