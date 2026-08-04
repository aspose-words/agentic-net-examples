using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class BatchWatermarkProcessor
{
    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the input directory exists; create it if missing.
        Directory.CreateDirectory(inputDir);
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // If the input folder is empty, create a sample DOCX file to process.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            string samplePath = Path.Combine(inputDir, "Sample.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document for watermark processing.");
            sampleDoc.Save(samplePath);
        }

        // Process each DOCX file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Configure a semi‑transparent text watermark.
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true // Ensure the watermark is semi‑transparent.
            };

            // Apply the watermark to the document.
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Determine the output file path and save the modified document.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Validate that the file was saved successfully.
            if (!File.Exists(outputPath))
                throw new IOException($"Failed to save watermarked document: {outputPath}");
        }
    }
}
