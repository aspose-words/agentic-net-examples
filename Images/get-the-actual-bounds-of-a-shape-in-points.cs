using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class GetShapeBounds
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape with a specific size.
        // Width = 200 points, Height = 150 points.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 150);

        // Position the shape relative to the left and top margins.
        shape.Left = 50;   // 50 points from the left margin.
        shape.Top = 75;    // 75 points from the top margin.

        // Retrieve the actual bounds of the shape in points.
        RectangleF boundsInPoints = shape.BoundsInPoints;

        // Output the bounds to the console.
        Console.WriteLine($"Shape Bounds (points): X={boundsInPoints.X}, Y={boundsInPoints.Y}, " +
                          $"Width={boundsInPoints.Width}, Height={boundsInPoints.Height}");

        // Save the document (optional, demonstrates lifecycle rule usage).
        doc.Save("ShapeBounds.docx");
    }
}
