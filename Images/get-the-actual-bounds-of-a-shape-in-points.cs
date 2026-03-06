using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Insert a rectangle shape and set its position, size, and rotation.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.Left = 150;   // X position in points.
        shape.Top = 200;    // Y position in points.
        shape.Rotation = 30; // Rotate the shape to demonstrate actual bounds calculation.

        // Use ShapeRenderer (inherits from NodeRendererBase) to get the actual bounds,
        // which include the effect of rotation.
        ShapeRenderer renderer = new ShapeRenderer(shape);
        RectangleF actualBounds = renderer.BoundsInPoints;

        // Display the actual bounds.
        Console.WriteLine($"Actual bounds (including rotation): X={actualBounds.X}, Y={actualBounds.Y}, Width={actualBounds.Width}, Height={actualBounds.Height}");

        // Save the document (optional).
        doc.Save("ShapeBounds.docx");
    }
}
