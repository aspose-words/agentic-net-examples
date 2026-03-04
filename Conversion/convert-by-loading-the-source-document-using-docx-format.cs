using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the document using the DOCX format constructor.
        Document doc = new Document(sourcePath);

        // Output the detected original load format (should be Docx).
        Console.WriteLine($"Original load format: {doc.OriginalLoadFormat}");

        // Example: save the loaded document to a new file.
        string outputPath = @"C:\Docs\CopyDocument.docx";
        doc.Save(outputPath);
    }
}
