using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);   // Inline rectangle.
        builder.InsertShape(ShapeType.Ellipse, 80, 80);      // Inline ellipse.
        builder.InsertShape(ShapeType.CloudCallout, RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 150, 200, 100, WrapType.None); // Floating shape.

        // Define the uniform brand fill color.
        Color brandColor = Color.FromArgb(255, 0, 120, 215); // Example blue shade.

        // Apply the fill color to every shape in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            shape.FillColor = brandColor;
        }

        // Validation: ensure all shapes have the expected fill color.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.FillColor.ToArgb() != brandColor.ToArgb())
                throw new InvalidOperationException("A shape did not receive the correct fill color.");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformFillShapes.docx");
        doc.Save(outputPath);
    }
}
