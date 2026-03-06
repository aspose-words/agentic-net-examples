using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be opened.
        string docPath = @"C:\Docs\Sample.docx";

        // Open the existing Word document using the Document(string) constructor.
        Document doc = new Document(docPath);

        // Example usage: write the document's plain text to the console.
        Console.WriteLine(doc.GetText());
    }
}
