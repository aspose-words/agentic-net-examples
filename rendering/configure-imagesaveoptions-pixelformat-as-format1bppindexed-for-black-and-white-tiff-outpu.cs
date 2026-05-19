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
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will be saved as a black‑and‑white TIFF image.");

        // Configure image save options for 1‑bpp TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use a compression scheme suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "Document.BlackAndWhite.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success (no interactive input required).
        Console.WriteLine("TIFF saved to: " + outputPath);
    }
}
