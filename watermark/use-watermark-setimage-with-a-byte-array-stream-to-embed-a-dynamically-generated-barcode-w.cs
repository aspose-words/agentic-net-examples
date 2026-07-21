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

        // A minimal PNG image (1x1 pixel, transparent) represented as a byte array.
        // This stands in for a dynamically generated barcode image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WQAAAAASUVORK5CYII=");

        // Load the image bytes into a memory stream.
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Ensure the stream position is at the beginning.
            imageStream.Position = 0;

            // Configure optional image watermark appearance.
            ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
            {
                IsWashout = false, // No washout effect.
                Scale = 5           // Scale factor (optional).
            };

            // Apply the image as a watermark using the stream overload.
            doc.Watermark.SetImage(imageStream, imgOptions);
        }

        // Save the resulting document.
        const string outputFile = "WatermarkedBarcode.docx";
        doc.Save(outputFile);
    }
}
