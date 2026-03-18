using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertPlainTextContentControl
{
    static void Main()
    {
        // Load an existing DOCX document (replace with your actual file path)
        Document doc = new Document(@"C:\Input\Template.docx");

        // Create a DocumentBuilder attached to the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the start of the target bookmark.
        // The bookmark name should match the one defined in the template.
        const string bookmarkName = "TargetBookmark";
        bool bookmarkFound = builder.MoveToBookmark(bookmarkName);
        if (!bookmarkFound)
        {
            Console.WriteLine($"Bookmark '{bookmarkName}' not found.");
            return;
        }

        // Insert a plain‑text content control (StructuredDocumentTag) at the bookmark position.
        // SdtType.PlainText creates a plain text content control.
        StructuredDocumentTag plainTextControl = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Optionally set placeholder text for the content control.
        plainTextControl.Title = "MyPlainTextControl";
        plainTextControl.PlaceholderName = "Enter text here";

        // Save the modified document (replace with your desired output path)
        doc.Save(@"C:\Output\Result.docx");
    }
}
