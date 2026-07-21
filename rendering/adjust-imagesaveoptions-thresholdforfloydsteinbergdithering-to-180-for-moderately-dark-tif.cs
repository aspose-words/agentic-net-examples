using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text.
        builder.Writeln("Sample text for TIFF conversion.");

        // Insert a tiny 1x1 black PNG image (base64 encoded).
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK7cAAAAASUVORK5CYII=");
        builder.InsertImage(png);

        // Configure ImageSaveOptions for TIFF with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 180 // Moderately dark threshold.
        };

        // Save the document as a TIFF image.
        string outPath = Path.Combine(artifactsDir, "output.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new Exception("TIFF file was not created.");
    }
}
