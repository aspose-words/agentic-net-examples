using System;
using System.IO;
using System.Linq;
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

        // Insert a few sample shapes.
        // Inline rectangle.
        builder.InsertShape(ShapeType.Rectangle, 120, 60);
        // Inline ellipse.
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        // Floating shape with explicit positioning.
        Shape floatingShape = builder.InsertShape(
            ShapeType.Star, RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 150, 100, 100, WrapType.None);
        floatingShape.StrokeColor = Color.DarkGray; // Optional styling.

        // Apply a uniform fill color to every shape in the document.
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Set the fill to a solid LightBlue color.
            shape.FillColor = Color.LightBlue;
        }

        // Save the document.
        string outputPath = "UniformFillShapes.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to save the document to '{outputPath}'.");
    }
}
