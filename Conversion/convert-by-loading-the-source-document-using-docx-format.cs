using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the document. The constructor automatically detects the DOCX format.
        Document doc = new Document(sourcePath);

        // Example: output the detected original load format.
        Console.WriteLine($"Original load format: {doc.OriginalLoadFormat}");
    }
}
