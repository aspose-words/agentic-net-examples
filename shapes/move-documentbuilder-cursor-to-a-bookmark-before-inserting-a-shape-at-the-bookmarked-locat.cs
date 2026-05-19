using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Document start.");

        // Define a bookmark where the shape will be inserted.
        builder.StartBookmark("MyShapeBookmark");
        builder.Writeln("Bookmark placeholder.");
        builder.EndBookmark("MyShapeBookmark");

        // Move the builder's cursor to the start of the bookmark.
        if (!builder.MoveToBookmark("MyShapeBookmark"))
            throw new InvalidOperationException("Bookmark 'MyShapeBookmark' not found.");

        // Insert a rectangle shape at the bookmark location.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = Color.LightBlue;
        shape.Stroke.Color = Color.DarkBlue;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkShape.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to create the output document.", outputPath);
    }
}
