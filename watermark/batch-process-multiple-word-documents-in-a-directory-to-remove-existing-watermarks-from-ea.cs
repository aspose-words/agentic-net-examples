using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary input folder and populate it with sample documents that contain watermarks.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        Directory.CreateDirectory(inputFolder);

        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a text watermark to the document.
            doc.Watermark.SetText($"Sample Watermark {i}");

            // Save the document to the input folder.
            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            doc.Save(inputPath);
        }

        // Create an output folder where the processed documents will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(outputFolder);

        // Process each .docx file in the input folder: remove its watermark and save to the output folder.
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

        // Optional: verify that output files were created (no console output required).
        // The program ends here.
    }
}
