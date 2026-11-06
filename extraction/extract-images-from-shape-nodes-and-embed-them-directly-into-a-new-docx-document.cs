using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageExtractionAndEmbedding
{
    public static void Main()
    {
        // Create a sample source document with an embedded image.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // 1x1 pixel PNG image (transparent) encoded in Base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            sourceBuilder.InsertImage(imageStream);
        }

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedSource = new Document(sourcePath);

        // Find all shape nodes that contain images.
        var imageShapes = loadedSource.GetChildNodes(NodeType.Shape, true)
                                      .OfType<Shape>()
                                      .Where(s => s.HasImage)
                                      .ToList();

        if (!imageShapes.Any())
            throw new InvalidOperationException("No image shapes were found in the source document.");

        // Create a new destination document.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Embed each extracted image into the destination document.
        foreach (Shape shape in imageShapes)
        {
            byte[] extractedBytes = shape.ImageData.ImageBytes;
            using (MemoryStream extractedStream = new MemoryStream(extractedBytes))
            {
                destBuilder.InsertImage(extractedStream);
                // Add a line break between images for readability.
                destBuilder.Writeln();
            }
        }

        const string destPath = "extracted-images.docx";
        destDoc.Save(destPath);

        // Validate that the destination document contains the embedded images.
        Document validateDoc = new Document(destPath);
        var embeddedImages = validateDoc.GetChildNodes(NodeType.Shape, true)
                                        .OfType<Shape>()
                                        .Where(s => s.HasImage)
                                        .Count();

        if (embeddedImages == 0)
            throw new InvalidOperationException("No images were embedded into the destination document.");

        // Cleanup: optional removal of temporary files can be added here if desired.
    }
}
