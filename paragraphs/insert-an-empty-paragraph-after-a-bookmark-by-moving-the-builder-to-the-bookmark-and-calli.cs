using System;
using Aspose.Words;

namespace InsertParagraphAfterBookmark
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a bookmark with some text.
            builder.StartBookmark("MyBookmark");
            builder.Write("Text inside bookmark.");
            builder.EndBookmark("MyBookmark");

            // Move the builder cursor to the position just after the bookmark.
            // Parameters: bookmark name, isStart = false (end of bookmark), isAfter = true (after the end).
            builder.MoveToBookmark("MyBookmark", false, true);

            // Insert an empty paragraph at the current position.
            builder.InsertParagraph();

            // Save the document to the local file system.
            doc.Save("Output.docx");
        }
    }
}
