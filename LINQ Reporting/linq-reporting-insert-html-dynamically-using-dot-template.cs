using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load a DOCX template that contains a bookmark named "HtmlPlaceholder".
        // Inside the bookmark place a ReportingEngine tag, e.g. <<[data.HtmlContent]>>.
        Document doc = new Document("Template.docx");

        // Data source with an HTML fragment that we want to insert.
        var data = new
        {
            HtmlContent = "<h1 style='color:blue;'>Hello World</h1>" +
                          "<p>This is <b>HTML</b> inserted via ReportingEngine.</p>"
        };

        // Build the report – the tag <<[data.HtmlContent]>> will be replaced with the raw HTML string.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "data");

        // After the report is built, move to the bookmark and insert the HTML fragment.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToBookmark("HtmlPlaceholder");
        builder.InsertHtml(data.HtmlContent);

        // Save the final document.
        doc.Save("Result.docx");
    }
}
