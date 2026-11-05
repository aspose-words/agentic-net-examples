using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define source document file names.
        string[] sourceFiles = { "Doc1.docx", "Doc2.docx" };

        // Create sample documents with some content.
        CreateSampleDocument(sourceFiles[0], "Hello World in Document 1.");
        CreateSampleDocument(sourceFiles[1], "Sample text for Document 2.");

        // Batch process: clear the content of each document.
        foreach (string filePath in sourceFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Delete all content in the document's range.
            doc.Range.Delete();

            // Save the cleared document with a new name.
            string clearedPath = Path.GetFileNameWithoutExtension(filePath) + "_cleared.docx";
            doc.Save(clearedPath);
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
