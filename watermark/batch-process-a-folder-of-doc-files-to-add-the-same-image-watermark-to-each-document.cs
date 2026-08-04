using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        string imagePath = Path.Combine(baseDir, "watermark.png");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(baseDir);

        // Create a deterministic 1x1 pixel PNG image if it does not exist.
        if (!File.Exists(imagePath))
        {
            // PNG bytes for a transparent 1x1 pixel image.
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
        }

        // Create sample DOC files if the input folder is empty.
        if (Directory.GetFiles(inputDir, "*.doc").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                string samplePath = Path.Combine(inputDir, $"Sample{i}.doc");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOC file: add the image watermark and save to the output folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.doc"))
        {
            Document doc = new Document(filePath);

            // Apply the same image watermark to the document.
            // Use the overload that accepts a file path and ImageWatermarkOptions.
            // Passing null for options uses default settings.
            doc.Watermark.SetImage(imagePath, null);

            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
