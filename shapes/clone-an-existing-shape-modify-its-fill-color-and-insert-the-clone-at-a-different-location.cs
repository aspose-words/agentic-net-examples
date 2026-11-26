using System;
using System.Drawing;
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

        // Insert the original shape (a rectangle) and set its fill color.
        Shape originalShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        originalShape.FillColor = Color.LightGray;

        // Add a paragraph break to create a different location.
        builder.Writeln();

        // Clone the original shape.
        Shape clonedShape = (Shape)originalShape.Clone(true);

        // Modify the fill color of the cloned shape.
        clonedShape.FillColor = Color.Blue;

        // Insert the cloned shape at the new location.
        builder.InsertNode(clonedShape);

        // Save the document.
        string outputPath = "ClonedShape.docx";
        doc.Save(outputPath);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to save the document to '{outputPath}'.");
        }
    }
}
