using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "ColorLzw.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");
        builder.Writeln("Another line to increase page content.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // The ImagePixelFormat enum does not provide an 8‑bpp indexed value.
        // Therefore we keep the default pixel format (32‑bpp ARGB) and only set the compression.
        options.TiffCompression = TiffCompression.Lzw; // LZW compression.

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
