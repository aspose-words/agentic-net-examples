using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);

        // Insert an ellipse shape.
        builder.InsertShape(ShapeType.Ellipse, 100, 100);

        // Create a sample image file to insert.
        string sampleImagePath = "sample.png";
        CreateSampleImage(sampleImagePath);

        // Insert the image as a shape.
        builder.InsertImage(sampleImagePath);

        // Save the document (optional, just to have a file).
        string docPath = "ShapesDocument.docx";
        doc.Save(docPath);

        // Traverse all shapes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            string shapeTypeName = shape.ShapeType.ToString();
            string imageFileName = $"shape_{i}_{shapeTypeName}.png";

            // Render the shape to an image file.
            shape.GetShapeRenderer().Save(imageFileName, new ImageSaveOptions(SaveFormat.Png));

            // Validate that the image file was created.
            if (!File.Exists(imageFileName))
                throw new InvalidOperationException($"Failed to create image file: {imageFileName}");
        }
    }

    // Helper method to create a simple placeholder PNG image.
    private static void CreateSampleImage(string path)
    {
        // This is a 1x1 pixel PNG (transparent) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);

        // Ensure the image file exists.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample image file: {path}");
    }
}
