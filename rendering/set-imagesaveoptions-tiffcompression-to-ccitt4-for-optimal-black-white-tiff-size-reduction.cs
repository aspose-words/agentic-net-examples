using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");

        // Configure TIFF save options with CCITT4 compression and black‑white color mode
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt4;
        options.ImageColorMode = ImageColorMode.BlackAndWhite;

        // Save the document as a TIFF file
        string tiffPath = Path.Combine(outputDir, "SampleCcitt4.tiff");
        doc.Save(tiffPath, options);

        // Verify that the file was created
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Output file information
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF saved to: {tiffPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
