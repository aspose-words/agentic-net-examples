using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Paths for the two TIFF files.
        string uncompressedPath = Path.Combine(outputDir, "uncompressed.tiff");
        string compressedPath = Path.Combine(outputDir, "compressed.tiff");

        // Create a sample document with several pages of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            if (i < 10)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as an uncompressed TIFF.
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None
        };
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the same document as a TIFF using CCITT3 compression.
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3
        };
        doc.Save(compressedPath, compressedOptions);

        // Verify that the compressed file is smaller.
        long uncompressedSize = new FileInfo(uncompressedPath).Length;
        long compressedSize = new FileInfo(compressedPath).Length;

        Console.WriteLine($"Uncompressed TIFF size: {uncompressedSize} bytes");
        Console.WriteLine($"CCITT3 compressed TIFF size: {compressedSize} bytes");

        if (compressedSize >= uncompressedSize)
            throw new InvalidOperationException("Compressed TIFF is not smaller than the uncompressed version.");

        Console.WriteLine("Compression successful: the CCITT3 TIFF is smaller than the uncompressed TIFF.");
    }
}
