using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.Writeln("This document will be saved as a TIFF image with custom dithering.");

        // Insert a simple 1x1 pixel PNG image (black) from a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK7cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imageStream);
        }

        // Configure TIFF save options with Floyd‑Steinberg dithering and a higher threshold.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 180
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "SampleDocument.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("TIFF file saved successfully to: " + outputPath);
    }
}
