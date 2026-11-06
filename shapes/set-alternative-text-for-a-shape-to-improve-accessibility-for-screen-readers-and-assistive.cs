using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SetShapeAlternativeText
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Set alternative text for accessibility.
        shape.AlternativeText = "Important notice rectangle";

        // Save the document.
        string outputPath = "ShapeAlternativeText.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Optional: Verify that the shape's alternative text was set correctly.
        Shape loadedShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        if (loadedShape.AlternativeText != "Important notice rectangle")
            throw new InvalidOperationException("Alternative text was not set correctly.");
    }
}
