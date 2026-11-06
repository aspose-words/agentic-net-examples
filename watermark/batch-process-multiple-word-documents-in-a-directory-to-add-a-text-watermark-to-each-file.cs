using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "WatermarkBatchDemo");
        Directory.CreateDirectory(workDir);

        // Create a subfolder for input documents.
        string inputDir = Path.Combine(workDir, "Input");
        Directory.CreateDirectory(inputDir);

        // Create a subfolder for output documents.
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Generate sample Word documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i}");
            string inputPath = Path.Combine(inputDir, $"Sample{i}.docx");
            sampleDoc.Save(inputPath);
        }

        // Define the watermark text.
        string watermarkText = "CONFIDENTIAL";

        // Process each .docx file in the input directory.
        foreach (string inputFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputFile);

            // Apply the text watermark.
            doc.Watermark.SetText(watermarkText);

            // Save the watermarked document to the output directory.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(inputFile));
            doc.Save(outputPath);
        }

        // Simple validation: ensure each output file exists.
        foreach (string outputFile in Directory.GetFiles(outputDir, "*.docx"))
        {
            if (File.Exists(outputFile))
            {
                Console.WriteLine($"Watermarked file created: {outputFile}");
            }
        }
    }
}
