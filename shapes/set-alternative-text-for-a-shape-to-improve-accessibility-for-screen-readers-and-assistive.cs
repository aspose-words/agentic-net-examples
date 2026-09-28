using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeAlternativeTextExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        // Set alternative text for accessibility.
        shape.AlternativeText = "Blue rectangle used as a decorative element";

        // Save the document.
        string outputPath = "ShapeAltText.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");

        // Optional: Verify that the shape's alternative text was set correctly.
        Shape savedShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        if (savedShape.AlternativeText != "Blue rectangle used as a decorative element")
            throw new Exception("Alternative text was not set correctly on the shape.");
    }
}
