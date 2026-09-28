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
        builder.Writeln("Hello, Aspose.Words TIFF compression example.");

        // Set up image save options for TIFF with CCITT3 compression and 1bpp indexed pixel format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt3;               // Smallest file size compression
        options.PixelFormat = ImagePixelFormat.Format1bppIndexed;       // 1-bit per pixel

        // Save the document as a TIFF image.
        string outputPath = "output.tiff";
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Output the file size (optional verification).
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}
