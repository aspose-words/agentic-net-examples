using System;
using System.IO;
using Aspose.Words;
using SkiaSharp;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample documents
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create a sample image file to be used as watermark (1x1 red pixel PNG)
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "watermark.png");
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

        // Create sample DOC files
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(docsFolder, $"Doc{i}.doc");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Load the watermark image into an SKBitmap (required by the current Aspose.Words API)
        using SKBitmap watermarkBitmap = SKBitmap.Decode(imagePath);
        if (watermarkBitmap == null)
        {
            Console.WriteLine("Failed to load watermark image.");
            return;
        }

        // Process each DOC file and add the image watermark
        string[] docFiles = Directory.GetFiles(docsFolder, "*.doc");
        foreach (string filePath in docFiles)
        {
            Document document = new Document(filePath);
            document.Watermark.SetImage(watermarkBitmap);

            string outputPath = Path.Combine(docsFolder,
                Path.GetFileNameWithoutExtension(filePath) + "_watermarked.doc");
            document.Save(outputPath);

            // Simple validation: ensure the output file exists
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Watermark applied and saved: {outputPath}");
            }
            else
            {
                Console.WriteLine($"Failed to save watermarked document: {outputPath}");
            }
        }
    }
}
