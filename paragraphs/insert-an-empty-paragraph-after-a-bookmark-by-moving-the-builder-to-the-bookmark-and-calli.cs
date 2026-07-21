using System;
using Aspose.Words;

public class InsertParagraphAfterBookmark
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a bookmark with some text inside it.
        builder.StartBookmark("MyBookmark");
        builder.Write("This is the bookmark content.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's cursor to the position just after the bookmark end.
        // Parameters: bookmark name, isStart = false (end of bookmark), isAfter = true (after the end node).
        builder.MoveToBookmark("MyBookmark", false, true);

        // Insert an empty paragraph at the current cursor position.
        builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("InsertedParagraph.docx");
    }
}
