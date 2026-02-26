using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains a bookmark named "HtmlPlaceholder"
        // where the dynamic HTML should be inserted.
        Document doc = new Document("Template.docx");

        // The HTML fragment that we want to inject into the document.
        string html = @"
            <h1>Hello <span style='color:blue;'>World</span></h1>
            <p>This is <b>dynamic</b> HTML inserted via Aspose.Words.</p>";

        // Use DocumentBuilder to position the cursor at the bookmark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToBookmark("HtmlPlaceholder");

        // Insert the HTML fragment. The InsertHtml method parses the HTML and
        // creates the corresponding Word objects (paragraphs, runs, styles, etc.).
        builder.InsertHtml(html);

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
