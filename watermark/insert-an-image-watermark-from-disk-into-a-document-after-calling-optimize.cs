using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders for output artifacts and temporary images.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        string imageDir = Path.Combine(Directory.GetCurrentDirectory(), "Images");

        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(imageDir);

        // Create a minimal 1x1 PNG image (red pixel) directly from a byte array.
        string imagePath = Path.Combine(imageDir, "watermark.png");
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
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document and add a line of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // Optimize the document for Word 2010 compatibility.
        doc.CompatibilityOptions.OptimizeFor(Aspose.Words.Settings.MsWordVersion.Word2010);

        // Define optional image watermark settings.
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 1,          // No scaling.
            IsWashout = false   // Preserve original colors.
        };

        // Apply the image watermark using the file path and options.
        doc.Watermark.SetImage(imagePath, options);

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "DocumentWithImageWatermark.docx");
        doc.Save(outputPath);

        // Simple validation: confirm that the file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? "Image watermark applied and document saved successfully."
            : "Failed to save the document.");
    }
}
