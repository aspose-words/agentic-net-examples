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

        // Save the document containing the rotated shape.
        const string docPath = "RotatedShape.docx";
        doc.Save(docPath);

        // Verify that the document file was created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The output document was not created.", docPath);

        // Retrieve the shape from the saved document to confirm the rotation value.
        Shape loadedShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (Math.Abs(loadedShape.Rotation - 45) > 0.001)
            throw new InvalidOperationException("The shape rotation was not set correctly.");

        // Render the shape to an image file to visually verify the rotation.
        const string imagePath = "RotatedShape.png";
        ShapeRenderer renderer = loadedShape.GetShapeRenderer();
        renderer.Save(imagePath, new ImageSaveOptions(SaveFormat.Png));

        // Verify that the rendered image file was created.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("The rendered shape image was not created.", imagePath);
    }
}
