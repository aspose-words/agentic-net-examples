using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output directories relative to the executable location.
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

        // Create sample documents with watermarks.
        for (int i = 1; i <= 3; i++)
        {
            string inputPath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document doc = new Document();
            // Add a simple text watermark.
            doc.Watermark.SetText($"Sample Watermark {i}");
            doc.Save(inputPath);
        }

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Remove the watermark if one exists.
            if (doc.Watermark.Type != WatermarkType.None)
                doc.Watermark.Remove();

            // Save the cleaned document to the output directory, preserving the original file name.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputDir, fileName);
            doc.Save(outputPath);
        }

        // The program ends here. Output files are saved in the OutputDocs folder.
    }
}
