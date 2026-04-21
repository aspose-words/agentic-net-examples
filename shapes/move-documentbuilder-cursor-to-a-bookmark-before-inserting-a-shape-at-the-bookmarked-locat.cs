using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace BookmarkShapeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark with some placeholder text.
            const string bookmarkName = "MyBookmark";
            builder.StartBookmark(bookmarkName);
            builder.Write("Text inside the bookmark.");
            builder.EndBookmark(bookmarkName);

            // Move the builder's cursor to the start of the bookmark.
            bool moved = builder.MoveToBookmark(bookmarkName);
            if (!moved)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' was not found.");

            // Insert a rectangle shape at the bookmark location.
            // Width = 100 points, Height = 50 points.
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Save the document to the local file system.
            const string outputPath = "BookmarkShape.docx";
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException($"Failed to create the output file '{outputPath}'.");
        }
    }
}
