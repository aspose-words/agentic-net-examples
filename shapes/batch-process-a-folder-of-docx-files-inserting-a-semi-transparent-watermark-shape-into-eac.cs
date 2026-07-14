using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files in the input folder.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document #{i}.");
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each DOCX file: add a semi‑transparent text watermark.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Configure watermark options (semi‑transparent).
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                IsSemitrasparent = true,               // Semi‑transparent.
                FontSize = 48,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal
            };

            // Apply the watermark.
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Validate that the file was saved.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save watermark‑ed document: {outputPath}");
        }
    }
}
