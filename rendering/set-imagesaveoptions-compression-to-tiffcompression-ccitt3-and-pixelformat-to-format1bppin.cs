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
        builder.Writeln("This is a sample document for TIFF compression test.");

        // Prepare output directory and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Compressed.tiff");

        // Configure ImageSaveOptions: TIFF format, CCITT3 compression, 1bpp indexed pixel format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the TIFF file.");
        }

        // Output the file size for reference.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved to: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
