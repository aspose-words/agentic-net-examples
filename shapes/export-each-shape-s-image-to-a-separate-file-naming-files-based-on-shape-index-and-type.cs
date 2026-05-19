using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ShapeImages");
        Directory.CreateDirectory(outputDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);

        // Insert an ellipse shape.
        builder.InsertShape(ShapeType.Ellipse, 120, 120);

        // Insert a simple 1x1 PNG image shape using a byte array (no System.Drawing dependency).
        // This is a transparent PNG encoded in base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        builder.InsertImage(pngBytes);

        // Optional: save the document for reference.
        string docPath = Path.Combine(outputDir, "SampleDocument.docx");
        doc.Save(docPath);

        // Traverse all shapes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .ToList();

        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];
            // Use the shape type name for the file name.
            string fileName = $"Shape_{i}_{shape.ShapeType}.png";
            string filePath = Path.Combine(outputDir, fileName);

            // Render the shape to a PNG image.
            ShapeRenderer renderer = shape.GetShapeRenderer();
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            renderer.Save(filePath, options);

            // Validate that the image file was created.
            if (!File.Exists(filePath))
                throw new Exception($"Failed to save image for shape index {i}.");

            Console.WriteLine($"Saved {filePath}");
        }
    }
}
