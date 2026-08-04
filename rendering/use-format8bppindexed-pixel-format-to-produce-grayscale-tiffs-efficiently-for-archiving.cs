using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Grayscale.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for grayscale TIFF rendering.");
        builder.Writeln("Second line of text.");

        // Configure TIFF save options for efficient grayscale output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Use a low‑bit pixel format to reduce file size.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Apply LZW compression for further size reduction.
            TiffCompression = TiffCompression.Lzw,
            // Set a reasonable resolution for archiving.
            Resolution = 300
        };

        // Save the document as a grayscale TIFF.
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optional: indicate success (no user interaction required).
        Console.WriteLine($"Grayscale TIFF saved to: {tiffPath}");
    }
}
