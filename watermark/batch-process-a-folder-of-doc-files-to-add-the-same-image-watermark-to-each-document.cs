using System;
using System.IO;
using Aspose.Words;
using SkiaSharp;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image for the watermark (1x1 transparent pixel)
        string imagePath = Path.Combine(baseDir, "watermark.png");
        if (!File.Exists(imagePath))
        {
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create sample DOC files if they do not already exist
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.doc");
            if (!File.Exists(docPath))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is sample document {i}.");
                doc.Save(docPath);
            }
        }

        // Load the watermark image into an SKBitmap once
        using SKBitmap watermarkBitmap = SKBitmap.Decode(imagePath);

        // Batch process: add the same image watermark to each DOC file
        foreach (string filePath in Directory.GetFiles(inputDir, "*.doc"))
        {
            Document doc = new Document(filePath);
            doc.Watermark.SetImage(watermarkBitmap);
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Simple validation output
        int processedCount = Directory.GetFiles(outputDir, "*.doc").Length;
        Console.WriteLine($"Watermarked {processedCount} document(s). Output folder: {outputDir}");
    }
}
