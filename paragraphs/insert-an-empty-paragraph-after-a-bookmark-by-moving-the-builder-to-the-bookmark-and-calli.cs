using System;
using Aspose.Words;

public class InsertEmptyParagraphAfterBookmark
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Paragraph before the bookmark.");

        // Insert a bookmark with some content.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Content inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Add more text after the bookmark (optional).
        builder.Writeln("Paragraph after the bookmark.");

        // Move the builder cursor to the position just after the bookmark end.
        // Parameters: bookmark name, isStart = false (end), isAfter = true (after the end).
        builder.MoveToBookmark("MyBookmark", false, true);

        // Insert an empty paragraph at the current cursor position.
        builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
