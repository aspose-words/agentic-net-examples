using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered as a black‑and‑white TIFF image.");

        // Prepare the folder for the output file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "BlackWhite.tiff");

        // Configure image save options for TIFF with 1‑bit indexed pixel format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use a CCITT compression scheme suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");

        // The program finishes automatically.
    }
}
