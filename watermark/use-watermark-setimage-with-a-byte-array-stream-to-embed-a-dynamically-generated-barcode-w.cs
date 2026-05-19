using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // A minimal 1x1 pixel PNG image (transparent) encoded in Base64.
        // This serves as a placeholder for a dynamically generated barcode image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Load the image bytes into a memory stream.
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            // Ensure the stream is positioned at the beginning.
            imageStream.Position = 0;

            // Configure optional image watermark settings.
            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                // Example: make the watermark more visible and larger.
                IsWashout = false,
                Scale = 5
            };

            // Apply the image watermark using the stream overload.
            doc.Watermark.SetImage(imageStream, options);
        }

        // Define the output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BarcodeWatermark.docx");

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
