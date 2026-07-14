using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create a sample document with multiple pages of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 10; i++)
        {
            builder.Writeln($"Page {i + 1}");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save without compression (TiffCompression.None).
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None
        };
        string uncompressedPath = Path.Combine(outputDir, "Uncompressed.tiff");
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save with CCITT3 compression.
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            // Use a binarization method suitable for CCITT compression.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 240
        };
        string compressedPath = Path.Combine(outputDir, "Compressed.tiff");
        doc.Save(compressedPath, compressedOptions);

        // Verify that both files exist.
        if (!File.Exists(uncompressedPath) || !File.Exists(compressedPath))
            throw new FileNotFoundException("One of the TIFF files was not created.");

        // Compare file sizes.
        long uncompressedSize = new FileInfo(uncompressedPath).Length;
        long compressedSize = new FileInfo(compressedPath).Length;

        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"CCITT3 compressed TIFF size: {compressedSize} bytes");

        // Ensure the compressed file is significantly smaller (e.g., less than 50% of the original).
        if (compressedSize >= uncompressedSize * 0.5)
            throw new InvalidOperationException("CCITT3 compression did not reduce the file size significantly.");
    }
}
