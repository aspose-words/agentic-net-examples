using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 PNG image file programmatically (no System.Drawing dependency).
        string sampleImagePath = "sample_image.png";
        // PNG data for a single blue pixel.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(sampleImagePath, pngData);

        // Create a new document and insert several shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);
        builder.Writeln();

        // Insert an ellipse shape.
        builder.InsertShape(ShapeType.Ellipse, 120, 120);
        builder.Writeln();

        // Insert an image shape using the sample image.
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        // Save the document.
        string docPath = "ShapesDocument.docx";
        doc.Save(docPath);

        // Traverse all shapes in the document and export each as an image file.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            string fileName = $"shape_{i}_{shape.ShapeType}.png";

            // Render the shape to a PNG image.
            shape.GetShapeRenderer().Save(fileName, new ImageSaveOptions(SaveFormat.Png));

            // Validate that the image file was created.
            if (!File.Exists(fileName))
                throw new Exception($"Failed to save shape image: {fileName}");
        }

        // Clean up the temporary sample image file.
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }
}
