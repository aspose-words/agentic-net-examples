using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class that will be used by the ReportingEngine.
    public class ReportData
    {
        public string Title { get; set; }          // Example plain text field.
        public string HtmlContent { get; set; }    // HTML fragment to be inserted dynamically.
    }

    class Program
    {
        static void Main()
        {
            // Load the LINQ Reporting template (a .docx file that contains LINQ tags such as <<[ds.Title]>>).
            Document doc = new Document("Template.docx");

            // Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Dynamic Report",
                HtmlContent = "<h2 style=\"color:blue;\">Hello from HTML</h2><p>This paragraph is inserted from a Markdown‑generated HTML fragment.</p>"
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" must match the name used in the template tags (e.g., <<[ds.Title]>>).
            engine.BuildReport(doc, data, "ds");

            // After the report is built, insert the HTML fragment at a predefined bookmark.
            // The template should contain a bookmark named "HtmlPlaceholder" where the HTML will appear.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToBookmark("HtmlPlaceholder");
            // InsertHtml parses the HTML string and converts it to Word formatting.
            builder.InsertHtml(data.HtmlContent);

            // Save the final document.
            doc.Save("Result.docx");
        }
    }
}
