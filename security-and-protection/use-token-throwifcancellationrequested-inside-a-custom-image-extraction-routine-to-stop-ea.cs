using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document and insert an in‑memory PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 1x1 pixel transparent PNG (base64 encoded).
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X5eUAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);

        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            // Insert the image into the document.
            builder.InsertImage(imgStream);
        }

        // Save the source document (optional, just to have a file on disk).
        const string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Set up cancellation support.
        CancellationTokenSource cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        try
        {
            // Run the custom image extraction routine.
            ExtractImages(doc, token, cts);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Image extraction was cancelled.");
        }
    }

    private static void ExtractImages(Document doc, CancellationToken token, CancellationTokenSource cts)
    {
        // Retrieve all shapes that may contain images.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Throw if cancellation was requested before processing the next shape.
            token.ThrowIfCancellationRequested();

            if (shape.HasImage)
            {
                string imageFile = $"Image_{imageIndex}.png";
                shape.ImageData.Save(imageFile);
                Console.WriteLine($"Saved {imageFile}");

                // Cancel after the first image to demonstrate early termination.
                if (imageIndex == 0)
                {
                    cts.Cancel();
                }

                imageIndex++;
            }
        }
    }
}
