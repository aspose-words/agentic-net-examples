using System;
using System.Drawing;
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

        // Insert a rectangle shape (100 points wide, 50 points high).
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.StrokeColor = Color.Blue;
        shape.FillColor = Color.LightGray;

        // Force layout so that actual bounds can be calculated.
        doc.UpdatePageLayout();

        // Retrieve the actual bounds of the shape.
        RectangleF bounds = shape.Bounds;

        // Log the coordinate points.
        Console.WriteLine("Shape Actual Bounds:");
        Console.WriteLine($"X: {bounds.X}");
        Console.WriteLine($"Y: {bounds.Y}");
        Console.WriteLine($"Width: {bounds.Width}");
        Console.WriteLine($"Height: {bounds.Height}");

        // Save the document.
        string outputPath = "ShapeBounds.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to save the document to '{outputPath}'.");

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
