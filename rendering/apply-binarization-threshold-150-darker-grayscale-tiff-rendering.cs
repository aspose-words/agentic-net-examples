using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        string imageDir = Path.Combine(Environment.CurrentDirectory, "Images");
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(artifactsDir);

        string sampleImagePath = Path.Combine(imageDir, "SampleImage.png");
        if (!File.Exists(sampleImagePath))
        {
            // 1x1 pixel PNG (transparent)
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=");
            File.WriteAllBytes(sampleImagePath, pngBytes);
        }

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");
        builder.InsertImage(sampleImagePath);

        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150
        };

        doc.Save(Path.Combine(artifactsDir, "DarkerGrayscale.tiff"), tiffOptions);
    }
}
