using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple PNG image (1x1 pixel, red) as a byte array.
        // PNG data taken from a minimal red pixel image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
        string imagePath = Path.Combine(artifactsDir, "RedPixel.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a blank Word document.
        Document doc = new Document();

        // Add some sample text so the watermarks are visible.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains both a text watermark and an image watermark.");

        // Apply a text watermark.
        doc.Watermark.SetText("Confidential");

        // Apply an image watermark on top of the text watermark.
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            // Scale the image to 30% of its original size.
            Scale = 0.3,
            // Disable washout to keep the image colors vivid.
            IsWashout = false
        };
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "CombinedWatermark.docx");
        doc.Save(outputPath);

        // Simple validation: ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
