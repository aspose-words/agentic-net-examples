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

        // Create a sample document with enough content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("Aspose.Words TIFF Compression Demo");
        builder.Writeln("This document will be saved twice:");
        builder.Writeln("- First without compression (TiffCompression.None)");
        builder.Writeln("- Then with CCITT3 compression (TiffCompression.Ccitt3)");
        // Add several paragraphs to increase the page size.
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"Line {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Save uncompressed TIFF.
        string uncompressedPath = Path.Combine(artifactsDir, "Uncompressed.tiff");
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None,
            // Ensure a reasonable resolution for size comparison.
            Resolution = 150
        };
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save CCITT3 compressed TIFF.
        string compressedPath = Path.Combine(artifactsDir, "Compressed_Ccitt3.tiff");
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            // Use the same resolution as the uncompressed version.
            Resolution = 150
        };
        doc.Save(compressedPath, compressedOptions);

        // Verify that both files exist.
        if (!File.Exists(uncompressedPath) || !File.Exists(compressedPath))
            throw new FileNotFoundException("One of the TIFF files was not created.");

        // Compare file sizes.
        long uncompressedSize = new FileInfo(uncompressedPath).Length;
        long compressedSize = new FileInfo(compressedPath).Length;

        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"CCITT3 compressed TIFF size: {compressedSize} bytes");

        // Expect a significant reduction (e.g., at least 30% smaller).
        if (compressedSize >= uncompressedSize * 0.7)
            throw new InvalidOperationException("CCITT3 compression did not reduce the file size significantly.");

        Console.WriteLine("Compression verification succeeded.");
    }
}
