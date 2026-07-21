using System;
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

        // Insert a simple cube shape.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
        shape.Name = "MyCube";

        // Set alternative text for accessibility.
        shape.AlternativeText = "Alt text for MyCube.";

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Shape_AltText.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Optional: Load the document again and confirm the alternative text.
        Document loadedDoc = new Document(outputPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape.AlternativeText != "Alt text for MyCube.")
            throw new InvalidOperationException("Alternative text was not set correctly.");
    }
}
