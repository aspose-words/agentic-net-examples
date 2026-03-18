using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Write("Bookmark content.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's insertion cursor to the start of the bookmark.
        // After this call the cursor is positioned just after the BookmarkStart node.
        builder.MoveToBookmark("MyBookmark");

        // Insert a rectangle shape at the bookmark location.
        // Width and height are measured in points (1 point = 1/72 inch).
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Save the resulting document.
        doc.Save("BookmarkShape.docx");
    }
}
