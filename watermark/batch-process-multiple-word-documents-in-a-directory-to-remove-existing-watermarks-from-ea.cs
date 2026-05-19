using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents with a text watermark.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");

            // Add a text watermark.
            sampleDoc.Watermark.SetText("Sample Watermark");

            // Save the document to the input folder.
            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            sampleDoc.Save(inputPath);
        }

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Remove the watermark if it exists.
            if (doc.Watermark.Type != WatermarkType.None)
            {
                doc.Watermark.Remove();
            }

            // Save the processed document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
