using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will be saved as a black‑and‑white TIFF image.");

        // Configure image save options for TIFF with 1‑bpp indexed pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages as black‑and‑white (1 bit per pixel).
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use CCITT4 compression which is suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4,
            // Set a reasonable resolution (dpi) for the output image.
            Resolution = 300
        };

        // Define the output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Document.BlackWhite.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, write a confirmation to the console (no user interaction required).
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
