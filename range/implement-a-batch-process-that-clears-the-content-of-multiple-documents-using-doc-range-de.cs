using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a few sample documents with some text.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(inputDir, $"Sample{i}.docx");
            CreateSampleDocument(filePath, $"This is the content of document {i}.");
        }

        // -----------------------------------------------------------------
        // Step 2: Batch process each document – clear its entire content.
        // -----------------------------------------------------------------
        foreach (string file in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(file);

            // Delete all characters in the whole‑document range.
            doc.Range.Delete();

            // Save the cleared document to the output folder.
            string fileName = Path.GetFileNameWithoutExtension(file);
            string outPath = Path.Combine(outputDir, $"{fileName}_Cleared.docx");
            doc.Save(outPath);
        }
    }

    // Helper method to create a simple document with a single paragraph of text.
    private static void CreateSampleDocument(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path);
    }
}
