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
        builder.Writeln("Another line to increase the file size for compression testing.");

        // Define output file paths.
        string uncompressedPath = "output_uncompressed.tiff";
        string compressedPath = "output_ccitt3.tiff";

        // Save the document as an uncompressed TIFF.
        ImageSaveOptions uncompressedOptions = new ImageSaveOptions(SaveFormat.Tiff);
        uncompressedOptions.TiffCompression = TiffCompression.None; // Explicitly no compression.
        doc.Save(uncompressedPath, uncompressedOptions);

        // Save the document as a TIFF using CCITT3 compression.
        ImageSaveOptions compressedOptions = new ImageSaveOptions(SaveFormat.Tiff);
        compressedOptions.TiffCompression = TiffCompression.Ccitt3;
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
        Console.WriteLine($"CCITT3 compressed TIFF size: {compressedSize} bytes");

        // Verify that compression reduced the file size significantly.
        if (compressedSize >= uncompressedSize)
            throw new InvalidOperationException("Compressed TIFF is not smaller than the uncompressed TIFF.");

        Console.WriteLine("Compression successful: the CCITT3 TIFF is smaller than the uncompressed TIFF.");
    }
}
