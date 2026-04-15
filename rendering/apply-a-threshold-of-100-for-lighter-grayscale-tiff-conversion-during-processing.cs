using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for lighter grayscale TIFF conversion.");

        // Configure TIFF save options with a threshold of 100.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,                     // Render as grayscale.
            TiffCompression = TiffCompression.Ccitt3,                     // Use CCITT3 compression.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 100                     // Apply threshold of 100.
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "GrayscaleThreshold.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
