using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Configure image save options for a color TIFF.
        // The ImagePixelFormat enum does not contain an 8‑bpp indexed value,
        // so we use the closest available color format (24‑bpp RGB) while keeping LZW compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PixelFormat = ImagePixelFormat.Format24BppRgb; // nearest supported color format
        options.TiffCompression = TiffCompression.Lzw;       // LZW compression

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(artifactsDir, "output.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
