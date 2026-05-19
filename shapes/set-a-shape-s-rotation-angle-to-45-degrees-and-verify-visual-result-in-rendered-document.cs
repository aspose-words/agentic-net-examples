using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Rendering; // Required for ShapeRenderer

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and insert a simple rectangle shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape of size 100x100 points.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);

        // Rotate the shape 45 degrees clockwise.
        shape.Rotation = 45;

        // Save the document.
        string docPath = Path.Combine(outputDir, "RotatedShape.docx");
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The document was not saved.");

        // Load the document and retrieve the shape to verify rotation.
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape == null)
            throw new InvalidOperationException("No shape found in the saved document.");

        if (Math.Abs(loadedShape.Rotation - 45) > 0.001)
            throw new InvalidOperationException($"Shape rotation is {loadedShape.Rotation}, expected 45.");

        // Render the rotated shape to an image (optional visual verification).
        string renderedPath = Path.Combine(outputDir, "RotatedShape.png");
        ShapeRenderer renderer = loadedShape.GetShapeRenderer();
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);
        renderer.Save(renderedPath, imgOptions);

        if (!File.Exists(renderedPath))
            throw new InvalidOperationException("The rendered shape image was not saved.");
    }
}
