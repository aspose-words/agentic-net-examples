using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline shape (a cube) with a size of 150x150 points.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
        shape.Name = "MyCube";

        // Set alternative text for accessibility.
        shape.AlternativeText = "Alt text for MyCube.";

        // Define the output file path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Shape_AltText.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optional: Verify that the alternative text was set correctly.
        Document loadedDoc = new Document(outputPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape.AlternativeText != "Alt text for MyCube.")
            throw new Exception("Alternative text was not set correctly.");
    }
}
