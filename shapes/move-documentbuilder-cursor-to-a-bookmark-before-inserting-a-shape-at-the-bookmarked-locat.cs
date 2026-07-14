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

        // Insert a bookmark named "MyBookmark" and some placeholder text.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This text is inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's cursor to the start of the bookmark.
        bool moved = builder.MoveToBookmark("MyBookmark");
        if (!moved)
            throw new Exception("Bookmark 'MyBookmark' was not found.");

        // Insert a rectangle shape at the bookmark location.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = System.Drawing.Color.LightBlue;
        shape.Stroke.Color = System.Drawing.Color.DarkBlue;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file at '{outputPath}'.");

        // Optional: indicate success.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
