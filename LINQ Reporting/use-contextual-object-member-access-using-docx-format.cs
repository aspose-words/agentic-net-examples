using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark named "MyBookmark" with some placeholder text.
        builder.StartBookmark("MyBookmark");
        builder.Write("Placeholder text");
        builder.EndBookmark("MyBookmark");

        // Move the cursor to the start of the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Overwrite the placeholder text with new content.
        builder.Write("Replaced text");

        // Save the document in DOCX format.
        doc.Save("Output.docx");
    }
}
