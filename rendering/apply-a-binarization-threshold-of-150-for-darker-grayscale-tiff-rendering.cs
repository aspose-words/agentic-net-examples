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
        builder.Writeln("Sample text for grayscale TIFF rendering with binarization threshold.");

        // Configure image save options for grayscale TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            TiffCompression = TiffCompression.Ccitt4
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the TIFF file.");
        }

        Console.WriteLine($"TIFF file successfully saved to: {outputPath}");
    }
}
