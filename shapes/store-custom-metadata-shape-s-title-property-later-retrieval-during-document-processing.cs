using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);

        // Store custom metadata in the shape's Title property.
        // Example format: key=value pairs separated by semicolons.
        shape.Title = "RecordId=42;Category=Invoice;Owner=AcmeCorp";

        // Save the document.
        string filePath = "ShapeWithMetadata.docx";
        doc.Save(filePath);

        // Load the document back.
        Document loadedDoc = new Document(filePath);

        // Retrieve the first shape in the document.
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Access the stored metadata.
        string metadata = loadedShape.Title;
        Console.WriteLine("Retrieved metadata: " + metadata);
    }
}
