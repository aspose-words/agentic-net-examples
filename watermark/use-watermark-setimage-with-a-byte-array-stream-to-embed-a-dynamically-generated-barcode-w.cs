using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // A minimal PNG image (1x1 pixel, black) represented as a byte array.
        // This serves as a placeholder for a dynamically generated barcode image.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52, // IHDR chunk
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01, // width=1, height=1
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
            0x00,0x00,0x00,0x0A,0x49,0x44,0x41,0x54,
            0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,0x00,
            0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };

        // Load the image bytes into a memory stream.
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Ensure the stream is positioned at the beginning.
            imageStream.Position = 0;

            // Configure watermark options (optional).
            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                // Example: make the watermark more visible.
                IsWashout = false,
                Scale = 5
            };

            // Apply the image watermark using the stream.
            doc.Watermark.SetImage(imageStream, options);
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BarcodeWatermark.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation: ensure the file was created.
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
