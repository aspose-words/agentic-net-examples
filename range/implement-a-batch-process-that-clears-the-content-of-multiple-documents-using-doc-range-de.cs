using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for the sample documents.
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(folder);

        // Names of the documents to process.
        string[] docNames = { "Doc1.docx", "Doc2.docx", "Doc3.docx" };
        List<string> docPaths = new List<string>();

        // Create sample documents with some text.
        foreach (string name in docNames)
        {
            string path = Path.Combine(folder, name);
            CreateSampleDocument(path, $"This is sample content for {name}");
            docPaths.Add(path);
        }

        // Batch process: clear the content of each document using Document.Range.Delete().
        foreach (string path in docPaths)
        {
            Document doc = new Document(path);
            doc.Range.Delete();               // Remove all nodes from the document.
            doc.Save(path);                    // Overwrite the original file.
        }

        // Simple verification: output the length of the remaining text (should be 0 or minimal).
        foreach (string path in docPaths)
        {
            Document doc = new Document(path);
            Console.WriteLine($"{Path.GetFileName(path)} text length after clear: {doc.Range.Text.Length}");
        }
    }

    // Helper method to create a simple document with a single paragraph of text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
