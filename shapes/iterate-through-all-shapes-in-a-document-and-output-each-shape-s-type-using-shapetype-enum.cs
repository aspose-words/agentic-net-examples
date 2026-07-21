using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several shapes of different types.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.Star, 60, 60);

        // Save the document to the local file system.
        string filePath = "ShapesIterate.docx";
        doc.Save(filePath);

        // Validate that the file was created.
        if (!File.Exists(filePath))
            throw new Exception($"Document was not saved to '{filePath}'.");

        // Load the saved document (optional, demonstrates load workflow).
        Document loadedDoc = new Document(filePath);

        // Retrieve all shape nodes in the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        // Output each shape's type to the console.
        foreach (Node node in shapeNodes)
        {
            if (node is Shape shape)
                Console.WriteLine(shape.ShapeType);
        }
    }
}
