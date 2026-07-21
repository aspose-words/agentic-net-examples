using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");

        // Create a simple 1x1 pixel PNG image if it does not already exist.
        if (!File.Exists(imagePath))
        {
            // PNG data for a single transparent pixel.
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark options: set scale and disable washout (full opacity).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // 50% of the original image size.
            IsWashout = false    // No washout effect; watermark appears with full opacity.
        };

        // Apply the image watermark using the image file and the configured options.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the document with the watermark.
        doc.Save(outputPath);
    }
}
