using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Folder to store the generated document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "ShapeTitleExample.docx");

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating cube shape.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);

        // Store custom metadata in the shape's Title property.
        const string customMetadata = "MyCustomMetadata";
        shape.Title = customMetadata;

        // Save the document.
        doc.Save(docPath);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Document was not saved to '{docPath}'.");

        // Load the document again and retrieve the shape.
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChildNodes(NodeType.Shape, true)[0];

        // Validate that the Title property contains the expected metadata.
        if (loadedShape.Title != customMetadata)
            throw new InvalidOperationException(
                $"Shape Title mismatch. Expected '{customMetadata}', but got '{loadedShape.Title}'.");

        // Indicate success.
        Console.WriteLine("Custom metadata stored in shape Title and retrieved successfully.");
    }
}
