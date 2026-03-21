using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // A tiny 1x1 red PNG image encoded in base64.
        const string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Image);
        using var imageStream = new MemoryStream(imageBytes);
        imageStream.Position = 0;

        // Create a new document and add text and the image.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with darker dithering.");
        builder.InsertImage(imageStream);

        // Configure TIFF save options with darker Floyd‑Steinberg dithering.
        var options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 200
        };

        // Ensure the output directory exists and save the TIFF.
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ArtifactsDir");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, "DarkerDithered.tiff");
        doc.Save(outputPath, options);

        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}
