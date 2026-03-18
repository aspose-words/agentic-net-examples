using Aspose.Words;
using System;

class InsertParagraphAfterBookmark
{
    static void Main()
    {
        // Load an existing document that contains the target bookmark.
        Document doc = new Document("Input.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the bookmark named "MyBookmark".
        // MoveToBookmark returns true if the bookmark exists.
        if (!builder.MoveToBookmark("MyBookmark"))
        {
            Console.WriteLine("Bookmark 'MyBookmark' not found.");
            return;
        }

        // Insert an empty paragraph immediately after the bookmark.
        builder.InsertParagraph();

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
