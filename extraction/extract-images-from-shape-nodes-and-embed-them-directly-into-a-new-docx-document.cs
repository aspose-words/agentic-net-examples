using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a source document that contains an image inside a shape.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // A tiny PNG image (1x1 pixel) encoded as Base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Insert the image as an inline shape.
            sourceBuilder.InsertImage(imageStream);
        }

        // Save the source document to the local file system.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document for extraction.
        Document loadedSource = new Document(sourcePath);

        // Create a new destination document where extracted images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Ensure the destination document has at least one paragraph to host images.
        destBuilder.Writeln("Extracted Images:");

        // Find all shape nodes that contain images.
        var shapeNodes = loadedSource.GetChildNodes(NodeType.Shape, true)
                                     .OfType<Shape>()
                                     .Where(s => s.HasImage);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            // Save the image data of the shape into a memory stream.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // Reset stream before reading.

                // Insert the image into the destination document.
                destBuilder.InsertImage(imgStream);
                destBuilder.Writeln(); // Add a line break after each image.
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted and inserted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the source document.");

        // Save the destination document containing the embedded images.
        const string destPath = "extracted_images.docx";
        destDoc.Save(destPath);

        // Additional validation: ensure the destination document now contains image shapes.
        var destImageShapes = destDoc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .Count(s => s.HasImage);
        if (destImageShapes == 0)
            throw new InvalidOperationException("The destination document does not contain any embedded images.");

        // Execution completed successfully.
    }
}
