using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ApplyUniformFillColor
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample shapes.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.Writeln(); // separate shapes with a paragraph break
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.Writeln();
        builder.InsertShape(ShapeType.Triangle, 120, 70);

        // Define the uniform fill color.
        Color uniformColor = Color.FromArgb(0, 112, 192); // Example branding blue.

        // Traverse all shapes in the document and apply the fill color.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            shape.FillColor = uniformColor;
        }

        // Save the document.
        string outputPath = "UniformFillShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
