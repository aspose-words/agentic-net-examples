using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertHtmlFromMhtml
{
    static void Main()
    {
        // Load the MHTML template. HtmlLoadOptions can be used to control loading behavior.
        Document template = new Document("Template.mht");

        // Create a DocumentBuilder to edit the loaded document.
        DocumentBuilder builder = new DocumentBuilder(template);

        // Move the cursor to a bookmark named "Content" where the HTML will be inserted.
        // If the bookmark does not exist, the builder will stay at the current position.
        builder.MoveToBookmark("Content");

        // HTML fragment to insert dynamically.
        string html = @"
            <h1 style='color:#1E90FF;'>Dynamic Title</h1>
            <p>This paragraph was inserted from an HTML string.</p>
            <ul>
                <li>Item 1</li>
                <li>Item 2</li>
            </ul>";

        // Insert the HTML using InsertHtml with options:
        // - UseBuilderFormatting: apply any formatting set on the builder as base formatting.
        // - RemoveLastEmptyParagraph: avoid an extra empty paragraph after the HTML block.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Save the resulting document. Here we save as DOCX, but you could also save as MHTML.
        template.Save("Result.docx", SaveFormat.Docx);
    }
}
