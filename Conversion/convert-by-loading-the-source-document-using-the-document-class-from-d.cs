using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the DOCX document using the Document constructor that accepts a file name.
        Document doc = new Document(sourcePath);

        // Example: output the total number of characters in the document.
        Console.WriteLine($"Document loaded. Character count: {doc.GetText().Length}");

        // Optional: save the loaded document to a new location.
        string outputPath = @"C:\Docs\CopyDocument.docx";
        doc.Save(outputPath);
    }
}
