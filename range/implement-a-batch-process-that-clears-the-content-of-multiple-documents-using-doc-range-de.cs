using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample source documents.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), "First document content.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), "Second document content.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc3.docx"), "Third document content.");

        // Process each document: load, clear its entire range, and save the cleared version.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Delete all characters in the document's range, effectively clearing its content.
            doc.Range.Delete();

            // Determine the output file path (same file name in the output folder).
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));

            // Save the cleared document.
            doc.Save(outputPath);
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch clearing completed.");
    }

    // Helper method to create a simple document with specified text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
