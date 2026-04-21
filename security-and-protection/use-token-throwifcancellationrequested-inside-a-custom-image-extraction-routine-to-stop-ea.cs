using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Base64 encoded 1x1 pixel PNG image.
    private const string SamplePngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6XK9cAAAAASUVORK5CYII=";

    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample document and extracted images.
        string docPath = Path.Combine(outputDir, "Sample.docx");
        string imagesDir = Path.Combine(outputDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // Create a sample document that contains an image.
        CreateSampleDocument(docPath);

        // Set up a cancellation token source.
        using var cts = new CancellationTokenSource();

        try
        {
            // Extract images; the routine will cancel after the first image.
            ExtractImages(docPath, imagesDir, cts.Token);
        }
        catch (OperationCanceledException)
        {
            // Expected cancellation – continue gracefully.
        }

        // Verify that at least one image was extracted.
        if (Directory.GetFiles(imagesDir).Length == 0)
            throw new InvalidOperationException("No images were extracted.");

        // Program finishes without waiting for user input.
    }

    // Creates a simple DOCX file with a single embedded PNG image.
    private static void CreateSampleDocument(string filePath)
    {
        // Decode the PNG image.
        byte[] imageBytes = Convert.FromBase64String(SamplePngBase64);
        using var imageStream = new MemoryStream(imageBytes);

        // Build the document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with an embedded image:");
        builder.InsertImage(imageStream);

        // Save the document.
        doc.Save(filePath);
    }

    // Extracts all images from the specified document into the target folder.
    // The method checks the cancellation token after each image is processed.
    private static void ExtractImages(string docPath, string targetFolder, CancellationToken token)
    {
        var doc = new Document(docPath);

        // Find all Shape nodes that contain images.
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save the image to a file.
            string imageFile = Path.Combine(targetFolder, $"Image_{imageIndex}{shape.ImageData.ImageType.ToString().ToLower()}");
            shape.ImageData.Save(imageFile);
            imageIndex++;

            // Cancel after the first image to demonstrate early termination.
            if (imageIndex == 1)
                token.ThrowIfCancellationRequested();

            // Check for cancellation before processing the next image.
            token.ThrowIfCancellationRequested();
        }
    }
}
