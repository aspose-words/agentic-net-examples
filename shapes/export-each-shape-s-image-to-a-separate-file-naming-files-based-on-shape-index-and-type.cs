using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

namespace ShapeExportExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rectangle shape.
            builder.InsertShape(ShapeType.Rectangle, 100, 50);

            // Insert an ellipse shape.
            builder.InsertShape(ShapeType.Ellipse, 80, 80);

            // Create a tiny PNG image (1x1 pixel) from a byte array.
            // This avoids the need for System.Drawing.
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            string imagePath = Path.Combine(outputDir, "sample.png");
            File.WriteAllBytes(imagePath, pngBytes);

            // Insert the image as a shape.
            builder.InsertImage(imagePath);

            // Save the document (optional, just to have a reference file).
            string docPath = Path.Combine(outputDir, "Shapes.docx");
            doc.Save(docPath);

            // Export each shape's visual representation to a separate PNG file.
            var shapeNodes = doc.GetChildNodes(NodeType.Shape, true)
                                .OfType<Shape>()
                                .ToList();

            for (int i = 0; i < shapeNodes.Count; i++)
            {
                Shape shape = shapeNodes[i];
                ShapeRenderer renderer = shape.GetShapeRenderer();

                string fileName = Path.Combine(
                    outputDir,
                    $"shape_{i}_{shape.ShapeType}.png");

                // Render the shape as PNG.
                renderer.Save(fileName, new ImageSaveOptions(SaveFormat.Png));
            }

            Console.WriteLine($"Exported {shapeNodes.Count} shape images to \"{outputDir}\".");
        }
    }
}
