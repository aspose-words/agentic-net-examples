using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineDemo
{
    // Simple POCO class that will serve as the data source for the report.
    public class ReportData
    {
        // This property contains Markdown formatted text.
        public string MarkdownText { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare the data source.
            var data = new ReportData
            {
                MarkdownText = "# Sample Report\n\n" +
                               "This is a **bold** statement.\n\n" +
                               "* Item 1\n" +
                               "* Item 2\n" +
                               "* Item 3\n"
            };

            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert a template tag that references the MarkdownText property.
            // The ':markdown' format tells the ReportingEngine to interpret the replacement as Markdown.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("<<[ds.MarkdownText]:markdown>>");

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The third argument ("ds") is the name used to reference the data source inside the template.
            engine.BuildReport(doc, data, "ds");

            // Save the resulting document.
            doc.Save("ReportWithMarkdown.docx");
        }
    }
}
