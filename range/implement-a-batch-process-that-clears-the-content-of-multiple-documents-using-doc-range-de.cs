using System;
using System.IO;
using Aspose.Words;

public class BatchClearRangeExample
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Sample document names.
        string[] docNames = { "Doc1.docx", "Doc2.docx", "Doc3.docx" };

        // Create sample source documents with some content.
        foreach (string name in docNames)
        {
            string inputPath = Path.Combine(inputFolder, name);
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of {name}.");
            doc.Save(inputPath);
        }

        // Batch process: load each document, clear its entire range, and save the result.
        foreach (string name in docNames)
        {
            string inputPath = Path.Combine(inputFolder, name);
            string outputPath = Path.Combine(outputFolder, name);

            // Load the document.
            Document doc = new Document(inputPath);

            // Delete all characters in the document's range, effectively clearing the content.
            doc.Range.Delete();

            // Save the cleared document.
            doc.Save(outputPath);
        }

        // Optional verification: ensure that the output documents are empty.
        foreach (string name in docNames)
        {
            string outputPath = Path.Combine(outputFolder, name);
            Document clearedDoc = new Document(outputPath);
            string text = clearedDoc.Range.Text.Trim();

            // The text should be empty after deletion.
            Console.WriteLine($"{name} cleared: {(string.IsNullOrEmpty(text) ? "Yes" : "No")}");
        }
    }
}
