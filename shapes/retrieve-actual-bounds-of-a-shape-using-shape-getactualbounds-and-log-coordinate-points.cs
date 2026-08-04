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
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left distance from page
            RelativeVerticalPosition.Page, 100,     // top distance from page
            200,                                    // width
            100,                                    // height
            WrapType.None);

        shape.StrokeColor = Color.Blue;

        // Retrieve the actual bounds of the shape.
        // The Shape class does not have GetActualBounds; use BoundsInPoints instead.
        RectangleF actualBounds = shape.BoundsInPoints;

        // Log the coordinate points.
        Console.WriteLine($"Actual Bounds: X={actualBounds.X}, Y={actualBounds.Y}, Width={actualBounds.Width}, Height={actualBounds.Height}");

        // Save the document to the local file system.
        doc.Save("ActualBounds.docx");
    }
}
