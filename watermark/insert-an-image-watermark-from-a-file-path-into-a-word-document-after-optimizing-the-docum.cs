using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Paths for the sample image and the resulting document.
        string imagePath = Path.Combine(outputDir, "watermark.png");
        string docPath = Path.Combine(outputDir, "DocumentWithImageWatermark.docx");

        // Write a minimal 1x1 transparent PNG to disk.
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

        // Create a new blank document.
        Document doc = new Document();

        // Add some text so the watermark can be seen.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // Optimize the document for Word 2010.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options (optional).
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            // Example: make the watermark opaque.
            IsWashout = false,
            // Example: scale the image (0 = auto).
            Scale = 0
        };

        // Insert the image watermark using the file path and options.
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the document.
        doc.Save(docPath);

        // Validate that the file was created.
        Console.WriteLine(File.Exists(docPath)
            ? $"Document saved successfully: {docPath}"
            : "Failed to save the document.");
    }
}
