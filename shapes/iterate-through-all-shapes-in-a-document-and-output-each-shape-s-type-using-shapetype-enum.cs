using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few different shapes.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.Line, 150, 0);
        builder.InsertShape(ShapeType.Star, 60, 60);

        // Save the document.
        string outputPath = "ShapesOutput.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");

        // Iterate through all shapes in the document and output their types.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int index = 1;
        foreach (Node node in shapeNodes)
        {
            if (node is Shape shape)
            {
                Console.WriteLine($"Shape {index}: {shape.ShapeType}");
                index++;
            }
        }
    }
}
