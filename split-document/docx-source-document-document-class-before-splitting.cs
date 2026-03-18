using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file that will be split later.
        string sourceFilePath = @"C:\Docs\Source.docx";

        // Load the document using the Document constructor that accepts a file name.
        Document sourceDocument = new Document(sourceFilePath);

        // Example verification – output the number of sections in the loaded document.
        Console.WriteLine($"Document loaded successfully. Sections count: {sourceDocument.Sections.Count}");
    }
}
