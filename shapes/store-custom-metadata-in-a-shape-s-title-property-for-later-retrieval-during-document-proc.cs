using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeTitleExample
{
    public static void Main()
    {
        // Define output folder and file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "ShapeWithTitle.docx");

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and store custom metadata in its Title property.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        string customTitle = "MyCustomMetadata_2023";
        shape.Title = customTitle;

        // Save the document.
        doc.Save(docPath);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Load the document and retrieve the shape.
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Validate that the Title property matches the stored metadata.
        if (loadedShape.Title != customTitle)
            throw new InvalidOperationException("The shape title does not match the expected value.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Shape title stored and retrieved successfully.");
    }
}
