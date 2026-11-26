using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Output TIFF file path.
        string outputPath = "binarized.tiff";

        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text for TIFF rendering with a binarization threshold of 150.");

        // Configure TIFF rendering options.
        // Aspose.Words does not expose a direct threshold property; using BlackAndWhite mode
        // together with CCITT4 compression produces a binary (binarized) image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.BlackAndWhite,
            TiffCompression = TiffCompression.Ccitt4,
            Resolution = 300
        };

        // Render the document to a TIFF file.
        doc.Save(outputPath, options);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        // Output basic information about the generated file.
        Console.WriteLine($"TIFF file saved successfully: {outputPath}");
        Console.WriteLine($"File size: {new FileInfo(outputPath).Length} bytes");
    }
}
