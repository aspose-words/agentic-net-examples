using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string filePath = @"C:\Docs\Sample.docx";

        // Load the document from the specified file.
        Document doc = new Document(filePath);

        // Example usage: output the text of the first paragraph.
        Console.WriteLine("Document loaded successfully.");
        Console.WriteLine("First paragraph text:");
        Console.WriteLine(doc.FirstSection.Body.FirstParagraph.GetText().Trim());
    }
}
