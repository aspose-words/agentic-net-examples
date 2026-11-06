using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the executable directory.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents with a text watermark.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");

            // Apply a text watermark using the native API.
            doc.Watermark.SetText("Sample Watermark");

            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(inputPath);
        }

        // Process each .docx file in the input folder to remove watermarks.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Remove any existing watermark.
            doc.Watermark.Remove();

            // Save the cleaned document to the output folder, preserving the original name.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);
            doc.Save(outputPath);
        }

        // Optional: Verify that output files were created.
        // (No console output as per requirements.)
    }
}
