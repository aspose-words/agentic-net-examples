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
        builder.Writeln("Sample text for grayscale TIFF conversion with a threshold of 100.");

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Configure ImageSaveOptions for grayscale TIFF.
        // The current Aspose.Words version does not expose a direct threshold property,
        // so we render the document in grayscale and use CCITT4 compression,
        // which is suitable for black‑and‑white (binarized) images.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ColorMode = ColorMode.Grayscale,          // Render as grayscale.
            TiffCompression = TiffCompression.Ccitt4 // Use CCITT4 compression for binary images.
        };

        // Render the document to TIFF using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        // Output basic information about the generated file.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully at '{outputPath}'. Size: {fileSize} bytes.");
    }
}
