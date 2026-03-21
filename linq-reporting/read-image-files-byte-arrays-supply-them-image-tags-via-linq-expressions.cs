#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageTagProcessor
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // Prepare a temporary image directory with a dummy PNG image.
        // -----------------------------------------------------------------
        string tempDir = Path.Combine(Path.GetTempPath(), "ImageTagDemo");
        Directory.CreateDirectory(tempDir);

        // Minimal 1x1 pixel PNG (transparent)
        byte[] dummyPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        string dummyImagePath = Path.Combine(tempDir, "SampleImage.png");
        File.WriteAllBytes(dummyImagePath, dummyPng);

        // -----------------------------------------------------------------
        // Load all image files from the directory and create a dictionary
        // where the key is the file name without extension and the value
        // is the image bytes.
        // -----------------------------------------------------------------
        var imageBytesByName = Directory.GetFiles(tempDir, "*.*")
            .Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".gif", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                f => Path.GetFileNameWithoutExtension(f),
                f => File.ReadAllBytes(f));

        // -----------------------------------------------------------------
        // Create a simple Word document with a shape that has a Title
        // matching the dummy image name.
        // -----------------------------------------------------------------
        Document doc = new Document();
        Shape shape = new Shape(doc, ShapeType.Image);
        shape.Title = "SampleImage"; // This will be used as the lookup key.
        using (var ms = new MemoryStream(dummyPng))
        {
            shape.ImageData.SetImage(ms);
        }

        // Shapes must be placed inside a paragraph.
        var paragraph = new Paragraph(doc);
        paragraph.AppendChild(shape);
        doc.FirstSection.Body.AppendChild(paragraph);

        // -----------------------------------------------------------------
        // Find all shapes that can contain an image.
        // -----------------------------------------------------------------
        var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
            .Cast<Shape>()
            .Where(s => s.HasImage);

        // -----------------------------------------------------------------
        // Replace each shape's image with the corresponding byte array from the dictionary.
        // -----------------------------------------------------------------
        foreach (Shape s in imageShapes)
        {
            string? key = s.Title;
            if (string.IsNullOrEmpty(key))
                continue;

            if (imageBytesByName.TryGetValue(key, out byte[] bytes))
            {
                using var ms = new MemoryStream(bytes);
                s.ImageData.SetImage(ms);
            }
        }

        // -----------------------------------------------------------------
        // Save the modified document to a temporary location.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(tempDir, "Result.docx");
        doc.Save(resultPath);
        Console.WriteLine($"Document saved to: {resultPath}");
    }
}
