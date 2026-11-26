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
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.Writeln();
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.Writeln();
        builder.InsertShape(ShapeType.Triangle, 100, 100);
        builder.Writeln();

        // Define the uniform fill color.
        Color uniformFill = Color.FromArgb(255, 0, 0); // Red

        // Apply the fill color to every shape in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            shape.FillColor = uniformFill;
        }

        // Save the document.
        string outputPath = "UniformFillShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to save the document to '{outputPath}'.");
        }
    }
}
