using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "SampleIndexedLzw.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words! This document will be saved as an indexed TIFF with LZW compression.");

        // Configure image save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The 8‑bpp indexed format is not available in this version of Aspose.Words.
            // Use the closest indexed format (1‑bpp) to preserve the intent of an indexed image.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Apply LZW compression (default for TIFF, but set explicitly).
            TiffCompression = TiffCompression.Lzw
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine("TIFF file created successfully at: " + outputPath);
    }
}
