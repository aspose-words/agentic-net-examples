using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Output document path
        string outputPath = "HyperlinkShape.docx";

        // URL to be applied to the shape
        string expectedUrl = "https://example.com";

        // Create a new document and a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and apply a hyperlink to it
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.HRef = expectedUrl; // Apply hyperlink

        // Save the document
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Reload the document and locate the shape
        Document loadedDoc = new Document(outputPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        if (shapes.Count == 0)
            throw new Exception("No shapes were found in the document.");

        Shape loadedShape = (Shape)shapes[0];

        // Verify that the shape contains the expected hyperlink
        if (loadedShape.HRef != expectedUrl)
            throw new Exception($"Hyperlink to '{expectedUrl}' was not found on the shape.");

        // Indicate success
        Console.WriteLine("Shape hyperlink applied and verified successfully.");
    }
}
