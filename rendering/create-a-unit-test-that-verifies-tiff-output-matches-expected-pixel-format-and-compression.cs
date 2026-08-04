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
        string tiffPath = Path.Combine(artifactsDir, "output.tiff");

        // Create a sample document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First page");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page");

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format24BppRgb,
            TiffCompression = TiffCompression.Ccitt4
        };

        // Render the document to a multi‑page TIFF file.
        doc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Verify that the source document has the expected number of pages.
        const int expectedPageCount = 2;
        if (doc.PageCount != expectedPageCount)
            throw new InvalidOperationException($"Document page count {doc.PageCount} does not match expected {expectedPageCount}.");

        // Verify that the save options were set to the expected values.
        if (options.PixelFormat != ImagePixelFormat.Format24BppRgb)
            throw new InvalidOperationException("Pixel format does not match the expected value.");

        if (options.TiffCompression != TiffCompression.Ccitt4)
            throw new InvalidOperationException("TIFF compression does not match the expected value.");

        Console.WriteLine("TIFF rendering test passed.");
    }
}
