using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the DOT template that contains a <<doc>> tag with the -sourceStyles switch,
            // e.g. <<doc [src.Html] -sourceStyles>>.
            Document template = new Document("Template.dot");

            // Prepare a data source that holds the HTML fragment to be inserted.
            var dataSource = new
            {
                // Any valid HTML can be placed here.
                Html = "<h2 style=\"color:blue;\">Dynamic Title</h2>" +
                       "<p>This is a <b>bold</b> paragraph inserted from LINQ Reporting.</p>"
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument is the name used in the template
            // (the "src" part of <<doc [src.Html] -sourceStyles>>).
            engine.BuildReport(template, dataSource, "src");

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
