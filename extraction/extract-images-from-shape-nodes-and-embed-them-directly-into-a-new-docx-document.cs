using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a source document containing an image shape.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // A tiny 1x1 PNG image encoded in Base64.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            sourceBuilder.InsertImage(imageStream);
        }

        sourceDoc.Save("source.docx");

        // Step 2: Load the source document.
        Document loadedSource = new Document("source.docx");

        // Step 3: Locate the first shape that contains an image.
        Shape imageShape = loadedSource
            .GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasImage);

        if (imageShape is null)
            throw new InvalidOperationException("No image shape found in the source document.");

        // Step 4: Extract the image bytes from the shape.
        byte[] extractedImageBytes = imageShape.ImageData.ImageBytes;
        if (extractedImageBytes == null || extractedImageBytes.Length == 0)
            throw new InvalidOperationException("Failed to extract image bytes.");

        // (Optional) Save the extracted image to verify extraction.
        File.WriteAllBytes("extracted-image.png", extractedImageBytes);
        if (!File.Exists("extracted-image.png"))
            throw new InvalidOperationException("Extracted image file was not created.");

        // Step 5: Create a new destination document and embed the extracted image.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        using (MemoryStream destImageStream = new MemoryStream(extractedImageBytes))
        {
            destBuilder.InsertImage(destImageStream);
        }

        destDoc.Save("output.docx");
        if (!File.Exists("output.docx"))
            throw new InvalidOperationException("Destination document was not created.");
    }
}
