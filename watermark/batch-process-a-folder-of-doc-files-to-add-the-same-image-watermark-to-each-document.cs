using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample documents and the watermark image.
        string sampleFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(sampleFolder);

        // Create a simple 1x1 PNG image to use as the watermark.
        string imagePath = Path.Combine(sampleFolder, "watermark.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Generate a few sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(sampleFolder, $"Doc{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Define watermark options (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the image to 50% of its original size.
            IsWashout = false    // Do not apply washout effect.
        };

        // Apply the image watermark to each DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(sampleFolder, "*.docx"))
        {
            Document doc = new Document(filePath);
            doc.Watermark.SetImage(imagePath, watermarkOptions);
            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }
    }
}
