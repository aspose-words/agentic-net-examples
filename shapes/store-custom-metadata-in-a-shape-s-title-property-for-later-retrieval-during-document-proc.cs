using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "ShapeMetadata.docx");

        // Create a new document and insert a rectangle shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Store custom metadata in the shape's Title property.
        string customTitle = "MyCustomMetadata123";
        shape.Title = customTitle;

        // Save the document.
        doc.Save(filePath);

        // Ensure the file was created.
        if (!File.Exists(filePath))
            throw new Exception("Failed to save the document.");

        // Load the document and retrieve the shape.
        Document loadedDoc = new Document(filePath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape == null)
            throw new Exception("Shape not found in the loaded document.");

        // Retrieve and verify the Title property.
        string retrievedTitle = loadedShape.Title;
        if (retrievedTitle != customTitle)
            throw new Exception($"Title mismatch. Expected: {customTitle}, Got: {retrievedTitle}");

        // Indicate success.
        Console.WriteLine("Shape title stored and retrieved successfully: " + retrievedTitle);
    }
}
