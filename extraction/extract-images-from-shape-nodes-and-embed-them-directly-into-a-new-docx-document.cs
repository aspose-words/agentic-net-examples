using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with an image shape.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // A tiny 1x1 PNG image encoded in base64.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            // Insert the image as a shape (inline by default).
            srcBuilder.InsertImage(imgStream);
        }

        // Save the source document to the local file system.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedDoc = new Document(sourcePath);

        // Prepare a new destination document where extracted images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Find all shape nodes that contain images.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Extract the raw image bytes and insert them into the destination document.
                byte[] imageBytes = shape.ImageData.ImageBytes;
                destBuilder.InsertImage(imageBytes);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No image shapes were found in the source document.");

        // Save the destination document containing the extracted images.
        const string destPath = "extracted_images.docx";
        destDoc.Save(destPath);

        // Optional verification: ensure the destination document actually contains shapes.
        int shapeCountInDest = destDoc.GetChildNodes(NodeType.Shape, true).Count;
        if (shapeCountInDest == 0)
            throw new InvalidOperationException("The destination document does not contain any image shapes.");

        // Execution completed successfully.
    }
}
