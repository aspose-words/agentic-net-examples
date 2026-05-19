using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a source document containing two images.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // 1x1 PNG (base64).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            sourceBuilder.InsertImage(ms);
        }

        sourceBuilder.Writeln(); // separate the images

        // 1x1 GIF (base64).
        byte[] gifBytes = Convert.FromBase64String(
            "R0lGODdhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==");
        using (MemoryStream ms = new MemoryStream(gifBytes))
        {
            sourceBuilder.InsertImage(ms);
        }

        // Save the source document to a local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the source document and locate image shapes.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .OfType<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

        if (imageShapes.Count == 0)
            throw new InvalidOperationException("No image shapes were found in the source document.");

        // -------------------------------------------------
        // 3. Create a new document that will hold the extracted images.
        // -------------------------------------------------
        Document resultDoc = new Document();
        // Ensure the document has the minimal required structure (Section → Body → Paragraph).
        resultDoc.EnsureMinimum();

        DocumentBuilder resultBuilder = new DocumentBuilder(resultDoc);
        // Position the builder at the end of the document (after the initial empty paragraph).
        resultBuilder.MoveToDocumentEnd();

        // -------------------------------------------------
        // 4. Insert each extracted image into the new document.
        // -------------------------------------------------
        foreach (Shape shape in imageShapes)
        {
            byte[] imageBytes = shape.ImageData.ImageBytes;
            if (imageBytes == null || imageBytes.Length == 0)
                continue; // safety check

            using (MemoryStream imageStream = new MemoryStream(imageBytes))
            {
                resultBuilder.InsertImage(imageStream);
            }

            // Add a paragraph break after each image for readability.
            resultBuilder.Writeln();
        }

        // -------------------------------------------------
        // 5. Save the result document.
        // -------------------------------------------------
        const string resultPath = "result.docx";
        resultDoc.Save(resultPath);

        // -------------------------------------------------
        // 6. Validate that images were embedded.
        // -------------------------------------------------
        int extractedImageCount = resultDoc.GetChildNodes(NodeType.Shape, true)
                                           .OfType<Shape>()
                                           .Count(s => s.HasImage);
        if (extractedImageCount == 0)
            throw new InvalidOperationException("No images were embedded into the result document.");

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(sourcePath);
        // File.Delete(resultPath);
    }
}
