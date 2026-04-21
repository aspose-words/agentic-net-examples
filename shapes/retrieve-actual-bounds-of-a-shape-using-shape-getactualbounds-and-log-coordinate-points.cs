using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,               // shape type
            RelativeHorizontalPosition.Page,   // horizontal reference
            100,                               // left position (points)
            RelativeVerticalPosition.Page,     // vertical reference
            100,                               // top position (points)
            200,                               // width (points)
            100,                               // height (points)
            WrapType.None);                    // no text wrapping

        // Retrieve the actual bounds of the shape using ShapeRenderer.
        ShapeRenderer renderer = new ShapeRenderer(shape);
        RectangleF actualBounds = renderer.BoundsInPoints;

        // Log the coordinate points.
        Console.WriteLine("Actual Bounds of the shape:");
        Console.WriteLine($"X = {actualBounds.X}");
        Console.WriteLine($"Y = {actualBounds.Y}");
        Console.WriteLine($"Width = {actualBounds.Width}");
        Console.WriteLine($"Height = {actualBounds.Height}");

        // Save the document to verify that the shape was inserted.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeActualBounds.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
