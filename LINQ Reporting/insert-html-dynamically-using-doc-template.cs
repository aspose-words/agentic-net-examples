using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains a bookmark named "HtmlPlace"
        Document doc = new Document("Template.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Dynamic HTML content to be inserted – this could come from any source
        string html = "<h2 align='center'>Dynamic Title</h2>" +
                      "<p>This is <b>bold</b> text with <i>italic</i> styling.</p>";

        // Position the builder at the bookmark where the HTML should be placed
        builder.MoveToBookmark("HtmlPlace");

        // Insert the HTML fragment into the document.
        // The InsertHtml method parses the HTML and converts it to Word formatting.
        builder.InsertHtml(html);

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
