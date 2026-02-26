using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data source containing an HTML fragment.
    public class ReportData
    {
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the RTF template that contains a bookmark named "HtmlPlaceholder".
            // The template should be placed in the same folder as the executable.
            Document doc = new Document("Template.rtf");

            // Prepare the data source with the HTML you want to insert.
            var data = new ReportData
            {
                HtmlContent = "<h2 style='color:#2E8B57;'>Dynamic Title</h2>" +
                              "<p>This paragraph is inserted <b>as HTML</b> at runtime.</p>"
            };

            // Use the LINQ ReportingEngine to populate the template.
            // The data source is referenced in the template by the name "ds".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // After the report is built, insert the HTML fragment at the bookmark.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToBookmark("HtmlPlaceholder");
            // InsertHtml parses the HTML and creates equivalent Word formatting.
            builder.InsertHtml(data.HtmlContent);

            // Save the final document.
            doc.Save("Result.docx");
        }
    }
}
