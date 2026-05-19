using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Determine base directory of the application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Create folders for input and output documents.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Generate sample source documents.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample content for document {i}.");
            string inputPath = Path.Combine(inputDir, $"Doc{i}.docx");
            sampleDoc.Save(inputPath);
        }

        // -----------------------------------------------------------------
        // Step 2: Batch process - clear the entire content of each document.
        // -----------------------------------------------------------------
        foreach (string inputPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Delete all characters in the document's range (clears content).
            doc.Range.Delete();

            // Save the cleared document to the output folder, preserving the file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(inputPath));
            doc.Save(outputPath);
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch clearing of documents completed.");
    }
}
