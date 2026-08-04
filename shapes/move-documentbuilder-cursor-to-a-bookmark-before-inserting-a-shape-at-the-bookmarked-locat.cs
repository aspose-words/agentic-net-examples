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

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bookmark named "MyBookmark".
            builder.StartBookmark("MyBookmark");
            builder.Writeln("This text is inside the bookmark.");
            builder.EndBookmark("MyBookmark");

            // Move the builder's cursor to the start of the bookmark.
            bool moved = builder.MoveToBookmark("MyBookmark");
            if (!moved)
                throw new InvalidOperationException("Bookmark not found.");

            // Insert a rectangle shape at the bookmark location.
            // Width = 100 points, Height = 50 points.
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Define the output file path (in the same folder as the executable).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkShape.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
