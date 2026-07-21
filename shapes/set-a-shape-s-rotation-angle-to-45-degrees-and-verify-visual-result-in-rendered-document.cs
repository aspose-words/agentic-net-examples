using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 100);
        // Rotate the shape 45 degrees clockwise.
        shape.Rotation = 45;

        // Define output file paths.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "RotatedShape.docx");
        string pngPath = Path.Combine(Directory.GetCurrentDirectory(), "RotatedShape.png");

        // Save the document.
        doc.Save(docPath);

        // Render the shape to an image file to visually verify the rotation.
        ShapeRenderer renderer = shape.GetShapeRenderer();
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        renderer.Save(pngPath, options);

        // Simple validation: ensure both files were created.
        if (!File.Exists(docPath))
            throw new Exception("Document file was not saved.");

        if (!File.Exists(pngPath))
            throw new Exception("Rendered shape image was not saved.");

        // Additional validation: check that the rotation property is set correctly.
        if (Math.Abs(shape.Rotation - 45) > 0.001)
            throw new Exception("Shape rotation was not set to 45 degrees.");
    }
}
