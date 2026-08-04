using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BatchWatermark
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents if the input folder is empty.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document {i}.");
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each .docx file in the input directory.
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string filePath in docFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Add a text watermark.
            doc.Watermark.SetText("Confidential");

            // Save the watermarked document to the output directory.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputDir, fileName);
            doc.Save(outputPath);

            Console.WriteLine($"Watermarked '{fileName}' and saved to OutputDocs.");
        }

        Console.WriteLine("Batch processing completed.");
    }
}
