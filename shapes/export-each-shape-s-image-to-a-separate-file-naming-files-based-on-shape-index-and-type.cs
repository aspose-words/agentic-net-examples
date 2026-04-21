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

        // Create a minimal PNG image (1x1 pixel) as a byte array.
        // This avoids the need for System.Drawing dependencies.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        // Create a new document and insert several shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 120, 60);

        // Insert an ellipse shape.
        builder.InsertShape(ShapeType.Ellipse, 80, 80);

        // Insert an image shape using the in‑memory PNG.
        Shape imageShape = new Shape(doc, ShapeType.Image);
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            imageShape.ImageData.SetImage(ms);
        }
        imageShape.Width = 100;
        imageShape.Height = 100;
        doc.FirstSection.Body.FirstParagraph.AppendChild(imageShape);

        // Save the document (optional, but fulfills the rule to save the document).
        string docPath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(docPath);

        // Export each shape to a separate image file.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            string fileName = $"{i}_{shape.ShapeType}.png";
            string filePath = Path.Combine(outputDir, fileName);

            if (shape.HasImage)
            {
                // Shape contains an image; save the raw image data.
                shape.ImageData.Save(filePath);
            }
            else
            {
                // Render the shape to an image.
                ShapeRenderer renderer = shape.GetShapeRenderer();
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
                renderer.Save(filePath, options);
            }

            // Validate that the file was created.
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create image file: {filePath}");
        }

        // Indicate completion.
        Console.WriteLine($"Exported {shapeNodes.Count} shape images to: {outputDir}");
    }
}
