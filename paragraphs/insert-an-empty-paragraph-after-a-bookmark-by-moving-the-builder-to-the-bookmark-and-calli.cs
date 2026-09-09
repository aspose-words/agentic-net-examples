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

            // Add a bookmark with some text inside.
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside the bookmark.");
            builder.EndBookmark("MyBookmark");

            // Move the builder's cursor to the bookmark.
            // This positions the cursor just after the start of the bookmark.
            builder.MoveToBookmark("MyBookmark");

            // Insert an empty paragraph at the current cursor position.
            builder.InsertParagraph();

            // Save the document to a file.
            doc.Save("Output.docx");
        }
    }
}
