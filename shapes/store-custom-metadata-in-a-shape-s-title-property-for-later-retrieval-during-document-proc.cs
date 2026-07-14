using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define an output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "ShapeMetadata.docx");

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and store custom metadata in its Title property.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.Title = "MyCustomMetadata";

        // Save the document containing the shape.
        doc.Save(docPath);

        // Load the document back from disk.
        Document loadedDoc = new Document(docPath);

        // Retrieve the first shape in the document.
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Verify that the Title property was preserved.
        if (loadedShape.Title != "MyCustomMetadata")
            throw new InvalidOperationException("Shape title metadata was not preserved.");

        // Output the retrieved title (no user interaction required).
        Console.WriteLine($"Retrieved shape title: {loadedShape.Title}");
    }
}
