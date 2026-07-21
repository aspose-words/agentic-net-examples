using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(imagesDir);

        // Create a sample document with two images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Small red dot PNG (1x1 pixel) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Insert the same image again to have more than one.
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Set up a cancellation token source that will be triggered after the first image.
        CancellationTokenSource cts = new CancellationTokenSource();

        try
        {
            ExtractImages(loadedDoc, imagesDir, cts);
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Image extraction was cancelled.");
        }

        // Verify that at least one image was saved.
        string[] extractedFiles = Directory.GetFiles(imagesDir);
        if (extractedFiles.Length == 0)
        {
            throw new InvalidOperationException("No images were extracted.");
        }

        Console.WriteLine($"Extraction completed. {extractedFiles.Length} image(s) saved to '{imagesDir}'.");
    }

    /// <summary>
    /// Extracts all images from the given document and saves them to the specified folder.
    /// The method checks the cancellation token on each iteration and aborts if cancellation is requested.
    /// </summary>
    private static void ExtractImages(Document doc, string outputFolder, CancellationTokenSource cts)
    {
        int imageIndex = 0;
        CancellationToken token = cts.Token;

        // Iterate over all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Throw if cancellation was requested.
            token.ThrowIfCancellationRequested();

            if (shape.HasImage)
            {
                // Determine the proper file extension for the image.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imagePath = Path.Combine(outputFolder, $"Image_{imageIndex}{extension}");

                // Save the image to a file.
                shape.ImageData.Save(imagePath);
                Console.WriteLine($"Saved image {imageIndex} to '{imagePath}'.");

                imageIndex++;

                // Cancel after the first image to demonstrate early termination.
                if (imageIndex == 1)
                {
                    cts.Cancel(); // Request cancellation for the next iteration.
                }
            }
        }
    }
}
