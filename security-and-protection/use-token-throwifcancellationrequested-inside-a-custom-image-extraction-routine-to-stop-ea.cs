using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Custom image extraction routine that respects cancellation.
    private static List<byte[]> ExtractImages(Document doc, CancellationToken token)
    {
        var images = new List<byte[]>();

        // Get all Shape nodes (including images) in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Allow cancellation before processing each shape.
            token.ThrowIfCancellationRequested();

            // Process only image shapes.
            if (shape.ShapeType == ShapeType.Image)
            {
                // Extract the image bytes.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    images.Add(ms.ToArray());
                }
            }
        }

        return images;
    }

    // Helper to create a simple 1x1 PNG image and return its stream.
    private static MemoryStream CreateSampleImage()
    {
        // Base64-encoded 1x1 pixel PNG (red color).
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        var ms = new MemoryStream(pngBytes);
        ms.Position = 0;
        return ms;
    }

    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first image.
        using (MemoryStream img1 = CreateSampleImage())
        {
            builder.InsertImage(img1);
        }

        // Insert second image.
        using (MemoryStream img2 = CreateSampleImage())
        {
            builder.InsertImage(img2);
        }

        // Save the document locally.
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Set up a cancellation token that will cancel after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 1 millisecond to ensure early cancellation.
            cts.CancelAfter(1);

            try
            {
                List<byte[]> extracted = ExtractImages(doc, cts.Token);
                Console.WriteLine($"Extracted {extracted.Count} images.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Image extraction was cancelled early.");
            }
        }

        // Verify that the output file exists.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The document was not saved as expected.");

        // End of program.
    }
}
