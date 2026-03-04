using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be processed.
        string docPath = "input.docx";

        // Load the document using the Document constructor (load rule).
        Document doc = new Document(docPath);

        // Extract plain‑text representation using PlainTextDocument (load rule).
        PlainTextDocument plainText = new PlainTextDocument(docPath);

        // Retrieve the concatenated text of the document.
        string extractedText = plainText.Text;

        // Output the extracted content.
        Console.WriteLine(extractedText);
    }
}
