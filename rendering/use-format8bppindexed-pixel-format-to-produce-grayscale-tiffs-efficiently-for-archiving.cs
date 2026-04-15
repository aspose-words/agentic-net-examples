using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for grayscale TIFF rendering.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("1234567890");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Use 1‑bit indexed pixel format for small file size.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Apply CCITT4 compression which is efficient for bi‑level images.
            TiffCompression = TiffCompression.Ccitt4,
            // Render each page as a separate frame in a multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(outputDir, "GrayscaleDocument.tiff");
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new Exception("Failed to create the TIFF file.");

        // Optionally, report success (no interactive input required).
        Console.WriteLine("Grayscale TIFF saved to: " + tiffPath);
    }
}
