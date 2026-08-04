using System;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string outputDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two images (red and green 1x1 PNGs)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First image (red square)
        using (MemoryStream redStream = CreateSamplePng(RedPngBase64))
        {
            builder.InsertImage(redStream);
        }

        // Second image (green square)
        using (MemoryStream greenStream = CreateSamplePng(GreenPngBase64))
        {
            builder.InsertImage(greenStream);
        }

        // Save the sample document
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // Load the document back
        Document loadedDoc = new Document(docPath);

        // Set up a cancellation token that will be triggered after the first image is saved
        CancellationTokenSource cts = new CancellationTokenSource();

        try
        {
            ExtractImages(loadedDoc, outputDir, cts.Token, cts);
        }
        catch (OperationCanceledException)
        {
            // Expected when cancellation is requested
            Console.WriteLine("Image extraction was cancelled as requested.");
        }

        // Verify that at least one image was extracted
        if (Directory.GetFiles(outputDir).Length == 0)
            throw new InvalidOperationException("No images were extracted.");
    }

    // Extracts all images from the document, respecting the cancellation token.
    private static void ExtractImages(Document doc, string folder, CancellationToken token, CancellationTokenSource cts)
    {
        int imageIndex = 0;

        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            // Throw if cancellation was requested before processing the next shape
            token.ThrowIfCancellationRequested();

            if (!shape.HasImage)
                continue;

            // Save the image to a file
            string imagePath = Path.Combine(folder,
                $"Image_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
            shape.ImageData.Save(imagePath);
            imageIndex++;

            // Cancel after the first image to demonstrate early termination
            if (imageIndex == 1)
                cts.Cancel();
        }
    }

    // Base64-encoded 1x1 PNG images (red and green)
    private const string RedPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z/C/HwAFgwJ/lKXK5wAAAABJRU5ErkJggg==";

    private const string GreenPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8DwHwAFgwJ/6cKX5wAAAABJRU5ErkJggg==";

    // Creates a MemoryStream from a base64-encoded PNG image.
    private static MemoryStream CreateSamplePng(string base64)
    {
        byte[] bytes = Convert.FromBase64String(base64);
        return new MemoryStream(bytes);
    }
}
