using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("This is a sample document rendered to a 24‑bpp color TIFF.");

        // Prepare the output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "Color24bpp.tiff");

        // Configure image save options for TIFF with 24‑bpp RGB pixel format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format24BppRgb,
            // Optional: set compression to None to keep the file size predictable.
            TiffCompression = TiffCompression.None
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
