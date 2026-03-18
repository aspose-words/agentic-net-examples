using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape with explicit size.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Position the shape on the page.
        shape.Left = 150;
        shape.Top = 200;
        // Apply a rotation to demonstrate that the actual bounds include rotation.
        shape.Rotation = 30;

        // -----------------------------------------------------------------
        // Retrieve the actual bounds of the shape.
        // BoundsInPoints returns the bounding rectangle of the shape as it
        // will be rendered on the page, taking rotation into account.
        // -----------------------------------------------------------------
        RectangleF actualBounds = shape.BoundsInPoints;
        Console.WriteLine($"Actual bounds (BoundsInPoints): X={actualBounds.X}, Y={actualBounds.Y}, Width={actualBounds.Width}, Height={actualBounds.Height}");

        // The same information can be obtained via a ShapeRenderer instance.
        ShapeRenderer renderer = shape.GetShapeRenderer();
        RectangleF rendererBounds = renderer.BoundsInPoints;
        Console.WriteLine($"Renderer bounds (BoundsInPoints): X={rendererBounds.X}, Y={rendererBounds.Y}, Width={rendererBounds.Width}, Height={rendererBounds.Height}");

        // Save the document (optional, demonstrates the full lifecycle).
        doc.Save("ShapeBounds.docx");
    }
}
