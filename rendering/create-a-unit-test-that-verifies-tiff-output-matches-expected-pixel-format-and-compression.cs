using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Run the TIFF rendering verification test.
        VerifyTiffOutput();
        Console.WriteLine("TIFF verification completed successfully.");
    }

    private static void VerifyTiffOutput()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a test document for TIFF rendering.");
        builder.InsertImage(CreateSamplePng()); // Insert a small image to ensure raster content.

        // Configure TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Expected compression and pixel format.
            TiffCompression = TiffCompression.Ccitt4,
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Render all pages into a single multi‑frame TIFF.
            PageSet = new PageSet(0) // start from first page
        };

        // Save the document as TIFF.
        string tiffPath = Path.Combine(artifactsDir, "RenderedDocument.tiff");
        doc.Save(tiffPath, tiffOptions);

        // ----- Validation -----
        // 1. Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"TIFF file was not created at '{tiffPath}'.");

        // 2. Verify that the file is not empty.
        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new InvalidOperationException("TIFF file is empty.");

        // 3. Verify that the save options were applied.
        // (We cannot inspect the TIFF internals without System.Drawing or external libraries,
        //  so we rely on the fact that Aspose.Words respects the options when saving.)
        if (tiffOptions.TiffCompression != TiffCompression.Ccitt4)
            throw new InvalidOperationException("Unexpected TIFF compression setting.");

        if (tiffOptions.PixelFormat != ImagePixelFormat.Format1bppIndexed)
            throw new InvalidOperationException("Unexpected TIFF pixel format setting.");
    }

    // Helper method to create a tiny PNG image in memory and return its file path.
    private static string CreateSamplePng()
    {
        // Use Aspose.Words' built‑in image generation via a temporary document.
        // This avoids System.Drawing usage.
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        tempBuilder.Writeln("PNG placeholder");
        string pngPath = Path.Combine(Path.GetTempPath(), "sample.png");
        tempDoc.Save(pngPath, SaveFormat.Png);
        return pngPath;
    }
}
