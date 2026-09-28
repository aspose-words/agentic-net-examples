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
        builder.Writeln("This is a sample document to demonstrate TIFF rendering with binarization.");

        // Configure image save options for TIFF.
        // The current Aspose.Words version does not expose a direct threshold property.
        // Use BlackAndWhite color mode with CCITT4 compression to obtain a binary (binarized) image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.BlackAndWhite,
            TiffCompression = TiffCompression.Ccitt4
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Save the document as a TIFF image with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
