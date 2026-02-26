using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Retrieve the first shape in the document (could be a Shape, GroupShape, etc.).
        ShapeBase shape = (ShapeBase)doc.GetChild(NodeType.Shape, 0, true);

        // Get the actual bounds of the shape in points.
        // This value does not include any rotation applied to the shape.
        RectangleF boundsInPoints = shape.BoundsInPoints;

        Console.WriteLine($"Shape BoundsInPoints: X={boundsInPoints.X}, Y={boundsInPoints.Y}, " +
                          $"Width={boundsInPoints.Width}, Height={boundsInPoints.Height}");

        // If you need the rendered bounds that take rotation into account,
        // use the ShapeRenderer (or OfficeMathRenderer for OfficeMath nodes).
        ShapeRenderer renderer = new ShapeRenderer(shape);
        RectangleF renderedBounds = renderer.BoundsInPoints;

        Console.WriteLine($"Rendered BoundsInPoints (with rotation): X={renderedBounds.X}, Y={renderedBounds.Y}, " +
                          $"Width={renderedBounds.Width}, Height={renderedBounds.Height}");

        // Save the document (optional, demonstrates the save lifecycle rule).
        doc.Save("Output.docx");
    }
}
