using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        shape.WrapType = WrapType.None; // Make it a floating shape.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100; // Position from the left edge of the page.
        shape.Top = 120;  // Position from the top edge of the page.

        // Retrieve the actual bounds of the shape after layout.
        // The BoundsInPoints property provides the shape's bounding rectangle in points.
        RectangleF actualBounds = shape.BoundsInPoints;

        // Log the coordinate points.
        Console.WriteLine("Actual Bounds:");
        Console.WriteLine($"  X      = {actualBounds.X}");
        Console.WriteLine($"  Y      = {actualBounds.Y}");
        Console.WriteLine($"  Width  = {actualBounds.Width}");
        Console.WriteLine($"  Height = {actualBounds.Height}");

        // Save the document to verify the shape is present.
        doc.Save("ActualBounds.docx");
    }
}
