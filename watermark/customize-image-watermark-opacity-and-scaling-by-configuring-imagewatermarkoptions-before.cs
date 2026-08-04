using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (1x1 pixel, red) as a byte array.
        // This avoids using System.Drawing APIs.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };

        // Write the image to a local file.
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark options: set scale and disable washout (full opacity).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5.0,          // Scale factor (e.g., 5 times the original size).
            IsWashout = false    // False means the watermark is not washed out (full opacity).
        };

        // Insert the image watermark using the file path and the configured options.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "Watermarked.docx");
        doc.Save(outputPath);
    }
}
