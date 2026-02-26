using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "input.docx";

        // Load the document using the DOCX format constructor.
        Document doc = new Document(sourcePath);

        // Example usage: output the document's plain text to the console.
        Console.WriteLine(doc.GetText());
    }
}
