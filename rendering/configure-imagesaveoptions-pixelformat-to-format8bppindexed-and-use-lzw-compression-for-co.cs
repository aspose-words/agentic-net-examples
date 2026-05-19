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
        builder.Writeln("This is a sample document rendered to a TIFF image.");

        // Prepare the folder for the output file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Sample.tiff");

        // Configure image save options:
        // - Save format: TIFF
        // - Pixel format: use the closest available indexed format (Format1bppIndexed)
        //   (Aspose.Words does not expose an 8‑bpp indexed format, so we use the nearest safe option).
        // - Compression: LZW (lossless compression for color TIFF).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            TiffCompression = TiffCompression.Lzw
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine($"TIFF image saved successfully to: {outputPath}");
    }
}
