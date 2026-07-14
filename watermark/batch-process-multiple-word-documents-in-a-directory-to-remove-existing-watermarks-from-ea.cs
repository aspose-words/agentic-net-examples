using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with watermarks.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            sampleDoc.Watermark.SetText($"Sample Watermark {i}");
            string samplePath = Path.Combine(inputDir, $"Doc{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each document: remove any existing watermark and save to output folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Remove watermark if present.
            if (doc.Watermark.Type != WatermarkType.None)
            {
                doc.Watermark.Remove();
            }

            // Save the processed document.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
