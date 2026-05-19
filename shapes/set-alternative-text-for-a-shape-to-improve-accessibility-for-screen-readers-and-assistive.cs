using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a cube shape and set its alternative text.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
        shape.AlternativeText = "Alt text for MyCube.";

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string docPath = Path.Combine(outputDir, "ShapeAltText.docx");
        doc.Save(docPath);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new Exception("The document was not saved correctly.");

        // Reload the document and verify the alternative text.
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape == null || loadedShape.AlternativeText != "Alt text for MyCube.")
            throw new Exception("Alternative text was not set correctly.");
    }
}
