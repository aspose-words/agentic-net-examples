using System;
using System.IO;
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
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.Cloud, 120, 70);

        // Define the uniform brand fill color.
        System.Drawing.Color brandColor = System.Drawing.Color.CornflowerBlue;

        // Traverse all shapes in the document and apply the brand fill color.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            shape.FillColor = brandColor; // Shortcut for solid fill.
        }

        // Save the document to the local file system.
        string outputPath = "UniformFillShapes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
