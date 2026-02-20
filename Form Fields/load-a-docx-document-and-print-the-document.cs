using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string docPath = "input.docx";

        // Load the Word document using the Document constructor that accepts a file name.
        Document doc = new Document(docPath);

        // Retrieve the full text of the document.
        string text = doc.GetText();

        // Print the document's text to the console.
        Console.WriteLine(text);
    }
}
