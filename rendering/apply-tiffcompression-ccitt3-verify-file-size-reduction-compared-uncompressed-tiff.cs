using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document with some text and a shape (ellipse) to make the file sizable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate TIFF compression.");
        // Insert an ellipse shape (200x200 points). This avoids the need for System.Drawing.
        builder.InsertShape(ShapeType.Ellipse, 200, 200);

        // Save the document as an uncompressed TIFF (no compression).
        string uncompressedPath = Path.Combine(artifactsDir, "Uncompressed.tiff");
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None,
            PageLayout = MultiPageLayout.TiffFrames()
        };
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the document as a TIFF using CCITT3 compression.
        string ccitt3Path = Path.Combine(artifactsDir, "Ccitt3.tiff");
        ImageSaveOptions ccitt3Options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            // CCITT compression works on 1‑bpp images; set binarization method accordingly.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            PageLayout = MultiPageLayout.TiffFrames()
        };
        doc.Save(ccitt3Path, ccitt3Options);

        // Compare file sizes.
        long uncompressedSize = new FileInfo(uncompressedPath).Length;
        long ccitt3Size = new FileInfo(ccitt3Path).Length;
        double reductionPercent = 100.0 * (uncompressedSize - ccitt3Size) / uncompressedSize;

        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"CCITT3 compressed TIFF size: {ccitt3Size} bytes");
        Console.WriteLine($"Size reduction: {reductionPercent:F2}%");

        // Simple verification: ensure the compressed file is significantly smaller.
        if (reductionPercent > 30) // arbitrary threshold for "significant"
        {
            Console.WriteLine("CCITT3 compression achieved a significant size reduction.");
        }
        else
        {
            Console.WriteLine("Compression did not achieve a significant size reduction.");
        }
    }
}
