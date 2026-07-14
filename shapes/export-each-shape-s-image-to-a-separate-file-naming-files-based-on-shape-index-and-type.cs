using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

public class ExportShapesExample
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // Create a new document and insert several shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 120, 60);

        // Insert an ellipse shape.
        builder.InsertShape(ShapeType.Ellipse, 80, 80);

        // Insert a text box shape.
        builder.InsertShape(ShapeType.TextBox, 150, 50);

        // Insert an image shape using a minimal in‑memory PNG (1×1 pixel).
        // This avoids the need for System.Drawing.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC");
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            // InsertImage returns the Shape that represents the image.
            builder.InsertImage(ms);
        }

        // Save the document (optional, just to have a reference file).
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        doc.Save(docPath);

        // Export each shape's visual representation to a separate PNG file.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.ShapeType != ShapeType.Group) // Skip group shapes (no appearance).
                        .ToList();

        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];
            // Render the shape to an image.
            ShapeRenderer renderer = shape.GetShapeRenderer();
            string fileName = $"shape_{i}_{shape.ShapeType}.png";
            string filePath = Path.Combine(imagesDir, fileName);
            renderer.Save(filePath, new ImageSaveOptions(SaveFormat.Png));

            // Validate that the file was created.
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to save shape image: {filePath}");
        }

        // Program ends automatically after Main finishes.
    }
}
