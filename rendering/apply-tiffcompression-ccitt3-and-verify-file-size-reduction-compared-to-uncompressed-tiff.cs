using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with enough content to span multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"This is page {i + 1} of the sample document.");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            // Insert a page break after each page except the last.
            if (i < 4)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as an uncompressed TIFF (no compression).
        string uncompressedPath = Path.Combine(outputDir, "Uncompressed.tiff");
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None
        };
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the same document as a TIFF using CCITT3 compression.
        string compressedPath = Path.Combine(outputDir, "Compressed_Ccitt3.tiff");
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3
        };
        doc.Save(compressedPath, compressedOptions);

        // Verify that both files were created.
        if (!File.Exists(uncompressedPath))
            throw new FileNotFoundException("Uncompressed TIFF was not created.", uncompressedPath);
        if (!File.Exists(compressedPath))
            throw new FileNotFoundException("Compressed TIFF was not created.", compressedPath);

        // Compare file sizes.
        long uncompressedSize = new FileInfo(uncompressedPath).Length;
        long compressedSize = new FileInfo(compressedPath).Length;

        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"Compressed (CCITT3) TIFF size: {compressedSize} bytes");

        // Ensure the compressed file is smaller than the uncompressed one.
        if (compressedSize >= uncompressedSize)
            throw new InvalidOperationException("Compressed TIFF is not smaller than the uncompressed TIFF.");

        // Optionally, check for a significant reduction (e.g., at least 30% smaller).
        double reductionRatio = (double)compressedSize / uncompressedSize;
        if (reductionRatio > 0.7)
            Console.WriteLine("Warning: Size reduction is less than 30%.");

        // Example completed successfully.
        Console.WriteLine("TIFF compression test completed.");
    }
}
