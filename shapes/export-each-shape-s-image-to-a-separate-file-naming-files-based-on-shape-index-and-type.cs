using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering; // Needed for ShapeRenderer
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few primitive shapes.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        builder.InsertShape(ShapeType.Ellipse, 80, 80);
        builder.InsertShape(ShapeType.Star, 70, 70);

        // Insert a simple 1x1 PNG image (transparent pixel) from a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            // InsertImage returns the Shape that contains the image.
            builder.InsertImage(ms);
        }

        // Save the document (optional, useful for inspection).
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleShapes.docx");
        doc.Save(docPath);

        // Prepare output directory for the exported shape images.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ShapeImages");
        Directory.CreateDirectory(outputDir);

        // Retrieve all Shape nodes from the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .ToList();

        // Export each shape to a separate PNG file.
        for (int i = 0; i < shapes.Count; i++)
        {
            Shape shape = shapes[i];

            // Render the shape to an image.
            ShapeRenderer renderer = shape.GetShapeRenderer();

            // Configure image saving options (PNG format).
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Build a file name that includes the shape index and its type.
            string fileName = $"shape_{i}_{shape.ShapeType}.png";
            string filePath = Path.Combine(outputDir, fileName);

            // Save the rendered image.
            renderer.Save(filePath, options);

            // Validate that the file was created.
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create image file: {filePath}");
        }

        // Indicate successful completion.
        Console.WriteLine("All shapes have been exported successfully.");
    }
}
