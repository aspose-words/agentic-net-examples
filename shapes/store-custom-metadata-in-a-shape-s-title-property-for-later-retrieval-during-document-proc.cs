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

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Store custom metadata in the shape's Title property.
        shape.Title = "CustomId:ABC123";

        // Save the document.
        string filePath = "ShapeMetadata.docx";
        doc.Save(filePath);

        // Verify that the file was saved.
        if (!File.Exists(filePath))
            throw new Exception("The document was not saved correctly.");

        // Load the document for later processing.
        Document loadedDoc = new Document(filePath);

        // Traverse all shapes and retrieve the Title property.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape s in shapeNodes)
        {
            if (!string.IsNullOrEmpty(s.Title))
            {
                Console.WriteLine($"Found shape with Title: {s.Title}");
            }
        }
    }
}
