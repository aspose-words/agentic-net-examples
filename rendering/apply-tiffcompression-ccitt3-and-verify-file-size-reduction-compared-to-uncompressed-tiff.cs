using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate TIFF compression with Aspose.Words.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        // Add enough content to make the TIFF size noticeable.
        for (int i = 0; i < 20; i++)
        {
            builder.Writeln($"Additional line {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Define file names.
        string uncompressedPath = "uncompressed.tiff";
        string compressedPath = "compressed.tiff";

        // Save the document as an uncompressed TIFF.
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None
        };
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the document as a TIFF using CCITT3 compression.
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

        // Output sizes for reference (optional).
        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"Compressed TIFF size:   {compressedSize} bytes");

        // Verify that compression reduced the file size significantly.
        if (compressedSize >= uncompressedSize)
            throw new InvalidOperationException("Compressed TIFF is not smaller than the uncompressed TIFF.");

        // Optionally, ensure a reasonable reduction (e.g., at least 10% smaller).
        double reductionPercent = 100.0 * (uncompressedSize - compressedSize) / uncompressedSize;
        if (reductionPercent < 10)
            throw new InvalidOperationException($"File size reduction ({reductionPercent:F2}%) is less than expected.");

        // Clean up resources (optional).
        // File.Delete(uncompressedPath);
        // File.Delete(compressedPath);
    }
}
