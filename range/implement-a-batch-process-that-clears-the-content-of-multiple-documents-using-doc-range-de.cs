using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents relative to the executable directory.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample source documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is the content of document {i}.");
            string inputPath = Path.Combine(inputDir, $"Doc{i}.docx");
            sampleDoc.Save(inputPath);
        }

        // Process each document: load, clear its entire content, and save to the output folder.
        foreach (string inputPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Delete all characters in the document's range, effectively clearing its content.
            doc.Range.Delete();

            // Save the cleared document to the output folder, preserving the original file name.
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputDir, fileName);
            doc.Save(outputPath);
        }
    }
}
