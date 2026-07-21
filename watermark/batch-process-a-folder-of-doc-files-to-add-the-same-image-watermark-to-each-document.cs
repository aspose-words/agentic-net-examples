using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Base directory for all sample data.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        string imagePath = Path.Combine(baseDir, "watermark.png");

        // Ensure base directory exists before creating subfolders.
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image for the watermark if it does not exist.
        if (!File.Exists(imagePath))
        {
            // 1x1 pixel transparent PNG (base64 encoded).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9W6vZc8AAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // Create sample DOCX files if the input folder is empty.
        if (Directory.GetFiles(inputDir, "*.doc*").Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                builder.Writeln("This document will receive an image watermark.");
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOC/DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.doc*"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Configure watermark options.
            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                Scale = 5,          // Example scale factor.
                IsWashout = false   // Keep original colors.
            };

            // Apply the image watermark.
            doc.Watermark.SetImage(imagePath, options);

            // Save the watermarked document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Validate that output files were created.
        foreach (string outFile in Directory.GetFiles(outputDir, "*.doc*"))
        {
            if (!File.Exists(outFile))
                throw new FileNotFoundException("Failed to create output file.", outFile);
        }
    }
}
