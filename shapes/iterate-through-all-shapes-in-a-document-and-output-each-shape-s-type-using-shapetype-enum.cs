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
        builder.InsertShape(ShapeType.Star, 60, 60);

        // Save the document to disk.
        string outputPath = "ShapesOutput.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Load the saved document (optional, demonstrates load workflow).
        Document loadedDoc = new Document(outputPath);

        // Retrieve all shape nodes in the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape and output its ShapeType.
        foreach (Node node in shapeNodes)
        {
            if (node is Shape shape)
            {
                Console.WriteLine(shape.ShapeType);
            }
        }
    }
}
