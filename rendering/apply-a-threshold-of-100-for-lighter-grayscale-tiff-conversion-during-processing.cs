using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 24;
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will be saved as a lighter grayscale TIFF image.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the document pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,

            // Use Floyd‑Steinberg dithering with a custom threshold of 100
            // to obtain a lighter grayscale result.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 100,

            // Use CCITT3 compression (suitable for 1‑bpp images).
            TiffCompression = TiffCompression.Ccitt3
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("TIFF file saved successfully: " + outputPath);
    }
}
