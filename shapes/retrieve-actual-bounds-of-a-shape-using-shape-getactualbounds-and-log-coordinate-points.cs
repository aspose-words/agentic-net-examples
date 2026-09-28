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

        // Insert a rectangle shape (inline) with specific size.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        shape.StrokeColor = Color.Blue;
        shape.FillColor = Color.LightGray;

        // Retrieve the actual bounds of the shape using the Bounds property.
        RectangleF bounds = shape.Bounds;

        // Log the coordinate points.
        Console.WriteLine("Shape Actual Bounds:");
        Console.WriteLine($"X (Left)   : {bounds.X}");
        Console.WriteLine($"Y (Top)    : {bounds.Y}");
        Console.WriteLine($"Width      : {bounds.Width}");
        Console.WriteLine($"Height     : {bounds.Height}");

        // Save the document to verify the shape is present.
        string outputPath = "ShapeBounds.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
    }
}
