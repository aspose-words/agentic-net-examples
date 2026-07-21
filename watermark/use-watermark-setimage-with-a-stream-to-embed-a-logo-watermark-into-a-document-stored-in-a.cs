using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Output file name.
        const string outputPath = "WatermarkedDocument.docx";

        // Create a blank document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // 1x1 transparent PNG encoded in Base64.
        const string pngBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6VQAAAAASUVORK5CYII=";

        // Convert Base64 string to a byte array.
        byte[] imageBytes = Convert.FromBase64String(pngBase64);

        // Apply the image watermark using a memory stream.
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            imageStream.Position = 0; // Ensure the stream is at the start.

            // Configure watermark appearance.
            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                Scale = 2.0,      // Make the watermark larger.
                IsWashout = false // Keep the image fully opaque.
            };

            // Set the image watermark from the stream.
            doc.Watermark.SetImage(imageStream, options);
        }

        // Save the document locally.
        doc.Save(outputPath);

        // Simple validation.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully: {Path.GetFullPath(outputPath)}"
            : "Failed to save the document.");
    }
}
