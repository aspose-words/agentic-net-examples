using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the executable location.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure clean directories.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with a text watermark.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            // Add a simple text watermark.
            sampleDoc.Watermark.SetText($"Sample Watermark {i}");
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Remove the watermark if it exists.
            if (doc.Watermark.Type != WatermarkType.None)
            {
                doc.Watermark.Remove();
            }

            // Save the processed document to the output directory, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: Verify that output files exist (no interactive output required).
        // The program ends here.
    }
}
