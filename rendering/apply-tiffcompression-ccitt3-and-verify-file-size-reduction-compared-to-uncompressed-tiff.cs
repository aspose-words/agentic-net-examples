using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several pages of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Page {i}");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            // Add a page break after each page except the last.
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define output folder (current directory).
        string outputFolder = Directory.GetCurrentDirectory();

        // Save the document as an uncompressed TIFF.
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.None
        };
        string uncompressedPath = Path.Combine(outputFolder, "uncompressed.tiff");
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the document as a TIFF using CCITT3 compression.
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3
        };
        string compressedPath = Path.Combine(outputFolder, "compressed.tiff");
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
            throw new InvalidOperationException("Compressed TIFF is not sufficiently smaller than the uncompressed version.");

        // Success message.
        Console.WriteLine("Compression verification passed.");
    }
}
