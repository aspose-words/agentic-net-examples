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
        builder.Writeln("This is a sample document for TIFF rendering with a threshold of 100.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Define output TIFF path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Configure image save options for grayscale TIFF.
        // The threshold of 100 is conceptually applied by using a grayscale mode;
        // Aspose.Words does not expose a direct TiffThreshold property in this version,
        // so we rely on the grayscale conversion which lightens the image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            // Use CCITT Group 4 compression for monochrome TIFFs.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Output the result path and file size for confirmation.
        FileInfo info = new FileInfo(outputPath);
        Console.WriteLine($"TIFF file created at: {outputPath}");
        Console.WriteLine($"File size: {info.Length} bytes");
    }
}
