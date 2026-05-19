using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    // Custom routine that extracts images from a document.
    // It checks the cancellation token on each iteration and aborts early if requested.
    private static void ExtractImages(Document doc, string outputFolder, CancellationTokenSource cts)
    {
        int imageIndex = 0;

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Iterate over all Shape nodes in the document (including those inside headers/footers).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Throw if cancellation was requested.
            cts.Token.ThrowIfCancellationRequested();

            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                // Build a file name for the extracted image.
                string extension = shape.ImageData.ImageType.ToString().ToLower(); // e.g., "png", "jpeg"
                string filePath = Path.Combine(outputFolder, $"Image_{imageIndex}.{extension}");

                // Save the image to disk.
                shape.ImageData.Save(filePath);
                Console.WriteLine($"Extracted: {filePath}");
                imageIndex++;

                // For demonstration, cancel after extracting the first image.
                if (imageIndex == 1)
                {
                    cts.Cancel();
                }
            }
        }

        Console.WriteLine($"Total images extracted before cancellation: {imageIndex}");
    }

    public static void Main()
    {
        // Folder paths used in the example.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imagesDir = Path.Combine(artifactsDir, "ExtractedImages");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with an embedded image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a tiny PNG image from a Base64 string (a 1x1 red pixel).
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string tempImagePath = Path.Combine(artifactsDir, "temp.png");
        File.WriteAllBytes(tempImagePath, pngBytes);

        // Insert the image into the document.
        builder.InsertImage(tempImagePath);
        builder.Writeln("Sample document with an image.");

        // Save the document locally.
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        doc.Save(docPath);
        Console.WriteLine($"Document saved to: {docPath}");

        // -----------------------------------------------------------------
        // 2. Load the document and extract images with cancellation support.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Set up a CancellationTokenSource.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            try
            {
                ExtractImages(loadedDoc, imagesDir, cts);
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Image extraction was cancelled.");
            }
        }

        // Clean up the temporary image file.
        if (File.Exists(tempImagePath))
        {
            File.Delete(tempImagePath);
        }

        // Indicate that the program has finished.
        Console.WriteLine("Execution completed.");
    }
}
