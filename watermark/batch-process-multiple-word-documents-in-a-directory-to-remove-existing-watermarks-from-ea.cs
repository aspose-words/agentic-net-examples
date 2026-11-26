using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for sample input documents and processed output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample Word documents with a text watermark.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a simple paragraph so the document is not empty.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");

            // Apply a text watermark using the native API.
            doc.Watermark.SetText("Sample Watermark");

            // Save the document to the input folder.
            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(inputPath);
        }

        // Batch process each .docx file in the input folder to remove watermarks.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Remove any existing watermark.
            doc.Watermark.Remove();

            // Save the cleaned document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: Verify that output files were created (no console output required).
        // The program ends here.
    }
}
