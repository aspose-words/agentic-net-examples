using System;
using System.IO;
using Aspose.Words;

public class WatermarkTool
{
    public static void Main(string[] args)
    {
        // Determine input directory and watermark text from command‑line arguments.
        string inputDir = args.Length > 0 ? args[0] : "InputDocs";
        string watermarkText = args.Length > 1 ? args[1] : "CONFIDENTIAL";

        // Ensure the directory exists.
        Directory.CreateDirectory(inputDir);

        // Find existing Word documents.
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");

        // If no documents are present, create a simple sample document.
        if (docFiles.Length == 0)
        {
            Document sample = new Document();
            // Add a paragraph so the document is not empty.
            sample.FirstSection.Body.FirstParagraph.AppendChild(new Run(sample, "Sample document content."));
            string samplePath = Path.Combine(inputDir, "Sample.docx");
            sample.Save(samplePath);
            docFiles = new[] { samplePath };
        }

        // Process each document: add the text watermark and save a new file.
        foreach (string filePath in docFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Apply a text watermark using the native API.
            doc.Watermark.SetText(watermarkText);

            // Build output file name.
            string outputPath = Path.Combine(
                inputDir,
                Path.GetFileNameWithoutExtension(filePath) + "_watermarked.docx");

            // Save the watermarked document.
            doc.Save(outputPath);
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Watermarking completed.");
    }
}
