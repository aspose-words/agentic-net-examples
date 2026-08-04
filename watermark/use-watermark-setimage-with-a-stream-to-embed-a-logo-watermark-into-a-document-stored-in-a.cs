using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1x1 pixel transparent PNG (base64 encoded)
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        // Simulate an Azure Blob stream.
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Ensure the stream is at the beginning.
            imageStream.Position = 0;

            // Create a new blank document.
            Document doc = new Document();

            // Add some content so the document is not empty.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document contains an image watermark.");

            // Configure watermark appearance.
            ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
            {
                Scale = 5,          // Scale factor.
                IsWashout = false   // Disable washout effect.
            };

            // Apply the image watermark from the stream.
            doc.Watermark.SetImage(imageStream, watermarkOptions);

            // Save the document.
            const string outputPath = "OutputWithWatermark.docx";
            doc.Save(outputPath);

            // Simple validation.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully: {outputPath}");
            }
            else
            {
                Console.WriteLine("Failed to save the document.");
            }
        }
    }
}
