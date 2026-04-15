using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string tiffPath = Path.Combine(artifactsDir, "output.tiff");

        // Create a simple multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,                     // Expected compression.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,            // Expected pixel format.
            PageLayout = MultiPageLayout.TiffFrames()                    // Multi‑page TIFF.
        };

        // Save the document as a TIFF image.
        doc.Save(tiffPath, options);

        // ----- Validation -----
        // 1. File must exist.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // 2. Options used for saving must match the expected values.
        if (options.TiffCompression != TiffCompression.Ccitt4)
            throw new Exception("TIFF compression does not match the expected value.");

        if (options.PixelFormat != ImagePixelFormat.Format1bppIndexed)
            throw new Exception("Pixel format does not match the expected value.");

        // 3. File size should be greater than zero.
        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new Exception("TIFF file is empty.");

        Console.WriteLine("TIFF output verification passed.");
    }
}
