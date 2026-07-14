using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample Word documents if they do not already exist.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            if (!File.Exists(samplePath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document {i}.");
                sampleDoc.Save(samplePath);
            }
        }

        // Iterate over each .docx file in the input directory and add a text watermark.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Configure watermark appearance.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Red,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply the watermark.
            doc.Watermark.SetText("CONFIDENTIAL", options);

            // Save the watermarked document to the output directory.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        Console.WriteLine("Watermark applied to all documents.");
    }
}
