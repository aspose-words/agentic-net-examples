using System;
using System.IO;
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

        // Insert sample shapes.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);
        builder.InsertShape(ShapeType.Ellipse, 100, 100);
        builder.InsertShape(ShapeType.Triangle, 120, 90);

        // Define the uniform fill color for branding.
        Color brandColor = Color.FromArgb(0, 112, 192); // Example brand blue.

        // Apply the fill color to every shape in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            shape.Fill.Color = brandColor;
        }

        // Save the document.
        string outputPath = "UniformFillShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to save the document to '{outputPath}'.");

        // Validate that each shape has the expected fill color.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.Fill.Color.ToArgb() != brandColor.ToArgb())
                throw new Exception("A shape does not have the expected fill color.");
        }
    }
}
