using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data class that will be used as the data source for the reporting engine.
    public class ReportData
    {
        public string TitleHtml { get; set; }
        public string BodyHtml { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains two bookmarks: "Title" and "Body".
            // The template can be created beforehand and placed in the same folder as the executable.
            Document template = new Document("Template.docx");

            // Prepare the data source. The properties contain HTML fragments that we want to insert.
            var data = new ReportData
            {
                TitleHtml = "<h1 style=\"color:#2E86C1;\">Dynamic Report Title</h1>",
                BodyHtml  = "<p>This paragraph is <b>bold</b> and this one is <i>italic</i>.</p>" +
                            "<ul><li>First item</li><li>Second item</li></ul>"
            };

            // Use the LINQ ReportingEngine to populate the template with the data source.
            // The data source name ("ds") can be referenced in the template if needed.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "ds");

            // After the report is built, insert the HTML fragments at the bookmark locations.
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert the title HTML at the "Title" bookmark.
            builder.MoveToBookmark("Title");
            builder.InsertHtml(data.TitleHtml);

            // Insert the body HTML at the "Body" bookmark.
            builder.MoveToBookmark("Body");
            builder.InsertHtml(data.BodyHtml);

            // Save the final document.
            template.Save("ReportOutput.docx");
        }
    }
}
