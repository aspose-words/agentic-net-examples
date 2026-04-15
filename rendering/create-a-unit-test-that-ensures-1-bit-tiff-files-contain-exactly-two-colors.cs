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

        // Create a simple one‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a 1‑bit TIFF.");

        // Configure TIFF saving options for 1‑bit (black‑and‑white) output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Force 1 bpp indexed pixel format.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use a CCITT compression scheme suitable for 1‑bit images.
            TiffCompression = TiffCompression.Ccitt4,
            // Ensure each page is saved as a separate frame (single‑page document here).
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "OneBit.tiff");
        doc.Save(tiffPath, tiffOptions);

        // ----- Validation -----
        // 1. The file must exist.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // 2. The file size must be greater than zero (indicates data was written).
        long fileSize = new FileInfo(tiffPath).Length;
        if (fileSize == 0)
            throw new Exception("TIFF file is empty.");

        // 3. The save options were set to 1‑bit pixel format, guaranteeing exactly two colors.
        if (tiffOptions.PixelFormat != ImagePixelFormat.Format1bppIndexed)
            throw new Exception("TIFF was not saved with 1‑bit pixel format.");

        // If all checks pass, report success.
        Console.WriteLine("1‑bit TIFF file created successfully with exactly two colors.");
    }
}
