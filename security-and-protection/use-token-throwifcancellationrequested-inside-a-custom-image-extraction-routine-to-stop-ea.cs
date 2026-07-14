using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Base64‑encoded 1×1 PNG images (red and green). They are tiny but sufficient for the demo.
    private const string RedPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8BQDwAF/AL+XKcAAAAASUVORK5CYII=";
    private const string GreenPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP4z8DAwAEAAf8C/6V5WQAAAABJRU5ErkJggg==";

    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        string imagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(imagesDir);

        // Create a sample document with two images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with sample images.");

        // First image (red square).
        using (MemoryStream redStream = new MemoryStream(Convert.FromBase64String(RedPngBase64)))
        {
            builder.InsertImage(redStream);
        }

        // Second image (green square).
        using (MemoryStream greenStream = new MemoryStream(Convert.FromBase64String(GreenPngBase64)))
        {
            builder.InsertImage(greenStream);
        }

        // Save the document.
        doc.Save(docPath);

        // Load the document back.
        Document loadedDoc = new Document(docPath);

        // Set up cancellation: the token can be cancelled from outside if needed.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            try
            {
                ExtractImages(loadedDoc, imagesDir, cts.Token);
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Image extraction was cancelled as requested.");
            }
        }

        // Verify how many images were saved.
        int savedCount = Directory.GetFiles(imagesDir, "*.png").Length;
        Console.WriteLine($"Number of images saved: {savedCount}");
    }

    // Extracts all images from the document, saving them to the target folder.
    // Throws if the supplied cancellation token is signaled.
    private static void ExtractImages(Document doc, string outputFolder, CancellationToken token)
    {
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Respect cancellation requests.
            token.ThrowIfCancellationRequested();

            if (shape.HasImage)
            {
                string fileName = Path.Combine(outputFolder, $"Image_{imageIndex}.png");
                shape.ImageData.Save(fileName);
                Console.WriteLine($"Saved image {imageIndex} to {fileName}");
                imageIndex++;

                // Demonstration: stop after the first image.
                // In a real scenario the caller would cancel the token source.
                if (imageIndex == 1)
                {
                    // No actual cancellation here; the method will simply continue.
                    // The token check at the start of the next iteration would stop the process
                    // if the caller had called cts.Cancel().
                }
            }
        }
    }
}
