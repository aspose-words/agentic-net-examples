using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ShapeActualBoundsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a floating rectangle shape.
            Shape shape = builder.InsertShape(
                ShapeType.Rectangle,               // Shape type.
                RelativeHorizontalPosition.Page,   // Position relative to the page horizontally.
                100,                               // Left position (points).
                RelativeVerticalPosition.Page,     // Position relative to the page vertically.
                100,                               // Top position (points).
                200,                               // Width (points).
                100,                               // Height (points).
                WrapType.None);                    // No text wrapping.

            // Retrieve the actual bounds of the shape.
            // Use the BoundsInPoints property which provides the location and size of the shape.
            RectangleF actualBounds = shape.BoundsInPoints;

            // Log the coordinate points.
            Console.WriteLine("Actual Bounds of the Shape:");
            Console.WriteLine($"X: {actualBounds.X}");
            Console.WriteLine($"Y: {actualBounds.Y}");
            Console.WriteLine($"Width: {actualBounds.Width}");
            Console.WriteLine($"Height: {actualBounds.Height}");

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeActualBounds.docx");
            doc.Save(outputPath);
        }
    }
}
