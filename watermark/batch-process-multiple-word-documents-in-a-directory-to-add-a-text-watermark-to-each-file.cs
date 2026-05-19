using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the folders exist.
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
                builder.Writeln($"This is the content of sample document {i}.");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Configure text watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = System.Drawing.Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true
            };

            // Apply the text watermark.
            doc.Watermark.SetText("CONFIDENTIAL", options);

            // Save the watermarked document to the output folder, preserving the original name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: simple verification that output files were created.
        foreach (string outFile in Directory.GetFiles(outputDir, "*.docx"))
        {
            Console.WriteLine($"Watermarked file created: {outFile}");
        }
    }
}
