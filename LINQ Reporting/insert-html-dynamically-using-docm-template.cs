using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertHtmlIntoDocmTemplate
{
    static void Main()
    {
        // Load the DOCM template.
        Document doc = new Document("Template.docm");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the HTML content to insert.
        string html = "<h2 align='center'>Dynamic Title</h2>" +
                      "<p>This paragraph was inserted <b>programmatically</b> from HTML.</p>";

        // Move the cursor to a bookmark named "ContentPlaceholder" in the template.
        // If the bookmark does not exist, this will throw; ensure the template contains it.
        builder.MoveToBookmark("ContentPlaceholder");

        // Insert the HTML. Use builder formatting as base and remove the extra empty paragraph
        // that Aspose.Words adds after block‑level HTML elements.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Save the result document.
        doc.Save("Result.docx");
    }
}
