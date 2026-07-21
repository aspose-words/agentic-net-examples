using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageExtractionAndEmbedding
{
    public static void Main()
    {
        // Create a source document with an embedded image.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Sample PNG image (1x1 pixel) encoded in Base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        // Insert the image into the source document.
        sourceBuilder.InsertImage(pngBytes);
        // Save the source document to the local file system.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document (demonstrates the load rule).
        Document loadedDoc = new Document(sourcePath);

        // Collect all shape nodes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        // Prepare the destination document where extracted images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data of the shape into a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                byte[] imageBytes = imageStream.ToArray();

                // Insert the extracted image into the destination document.
                destBuilder.InsertImage(imageBytes);
                // Add a line break after each image for readability.
                destBuilder.Writeln();
                extractedCount++;
            }
        }

        // Ensure that at least one image was extracted; otherwise throw.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the source document.");

        // Save the destination document containing the embedded images.
        const string destPath = "extracted_images.docx";
        destDoc.Save(destPath);
    }
}
