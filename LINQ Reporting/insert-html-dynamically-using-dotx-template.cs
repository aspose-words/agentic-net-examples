using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains a bookmark named "HtmlContent"
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to edit the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the bookmark where the HTML should be inserted
        builder.MoveToBookmark("HtmlContent");

        // Dynamic HTML to be inserted – this could come from any source (e.g., database, API)
        string html = @"
            <h2 style='color:blue;'>Dynamic Title</h2>
            <p>This is <b>bold</b> and <i>italic</i> text.</p>
            <ul>
                <li>Item 1</li>
                <li>Item 2</li>
            </ul>";

        // Insert the HTML fragment into the document at the current position
        builder.InsertHtml(html);

        // Save the populated document
        doc.Save("Result.docx");
    }
}
