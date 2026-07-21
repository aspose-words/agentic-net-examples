using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark named "MyBookmark" and add some text inside it.
        builder.StartBookmark("MyBookmark");
        builder.Write("This text is inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's cursor to the start of the bookmark.
        bool moved = builder.MoveToBookmark("MyBookmark");
        if (!moved)
            throw new InvalidOperationException("Bookmark 'MyBookmark' was not found.");

        // Insert a rectangle shape at the bookmark location.
        // Using InsertShape for simplicity; width and height are in points.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Optional: set a fill color to make the shape visible.
        shape.FillColor = System.Drawing.Color.LightBlue;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
