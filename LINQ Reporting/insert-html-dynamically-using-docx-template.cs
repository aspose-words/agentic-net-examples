using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlIntoTemplate
{
    static void Main()
    {
        // Load the DOCX template that contains a bookmark named "HtmlPlace".
        Document template = new Document("Template.docx");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move the cursor to the bookmark where the HTML should be inserted.
        builder.MoveToBookmark("HtmlPlace");

        // The HTML fragment to insert – this can be generated at runtime.
        string htmlFragment = "<h2 style=\"color:#2E86C1;\">Dynamic Title</h2>" +
                              "<p>This paragraph was inserted <b>programmatically</b> from HTML.</p>" +
                              "<ul><li>Item 1</li><li>Item 2</li></ul>";

        // Insert the HTML into the document at the current cursor position.
        // The overload without extra parameters uses the default formatting.
        builder.InsertHtml(htmlFragment);

        // Save the resulting document.
        template.Save("Result.docx", SaveFormat.Docx);
    }
}
