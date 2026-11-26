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

        // Insert a rectangle shape and set its size.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);

        // Store custom metadata in the shape's Title property.
        shape.Title = "CustomMetadata:ABC123";

        // Save the document to a temporary file.
        string filePath = Path.Combine(Path.GetTempPath(), "ShapeMetadata.docx");
        doc.Save(filePath);

        // Load the document for later processing.
        Document loadedDoc = new Document(filePath);

        // Traverse all shapes in the document and look for the stored metadata.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        bool metadataFound = false;

        foreach (Node node in shapeNodes)
        {
            if (node is Shape s && s.Title == "CustomMetadata:ABC123")
            {
                metadataFound = true;
                break;
            }
        }

        // Validate that the metadata was retrieved successfully.
        if (!metadataFound)
            throw new InvalidOperationException("The custom metadata was not found in any shape.");
    }
}
