using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputDocs";

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), "This is the first sample document.");
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), "Second document with different content.");

        // Batch process: clear content of each document and save the result.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Clear all content using the Range API.
            doc.Range.Delete();

            // Save the cleared document to the output folder.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);
            doc.Save(outputPath);
        }
    }

    // Helper method to create a simple document with given text.
    private static void CreateSampleDocument(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path);
    }
}
