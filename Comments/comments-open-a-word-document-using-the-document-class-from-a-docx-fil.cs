using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be opened.
        string docPath = "MyDir\\Sample.docx";

        // Open the Word document. The constructor automatically detects the file format.
        Document doc = new Document(docPath);

        // Example usage: write the document's plain text to the console.
        Console.WriteLine(doc.GetText());
    }
}
