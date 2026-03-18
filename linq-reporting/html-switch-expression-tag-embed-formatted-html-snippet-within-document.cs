using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedHtmlSnippet
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the HTML snippet to embed.
        // The snippet contains formatting (color, bold) that will be preserved.
        const string htmlSnippet = "<p style='color:red; text-align:center;'>" +
                                   "Hello <b>World</b> from <i>HTML</i>!" +
                                   "</p>";

        // Insert the HTML into the document. The InsertHtml method parses the HTML
        // and converts it to the corresponding WordML elements.
        builder.InsertHtml(htmlSnippet);

        // Optionally, set additional save options (e.g., pretty‑format the output HTML).
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            PrettyFormat = true
        };

        // Save the document to a .docx file.
        doc.Save("EmbeddedHtml.docx");

        // Also save as HTML to demonstrate that the inserted HTML is retained.
        doc.Save("EmbeddedHtml.html", saveOptions);
    }
}
