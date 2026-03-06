using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be processed.
        string docPath = "input.docx";

        // Load the Word document (uses Document(string) constructor).
        Document doc = new Document(docPath);

        // Extract plain‑text representation of the document (uses PlainTextDocument(string) constructor).
        PlainTextDocument plainText = new PlainTextDocument(docPath);

        // Retrieve the concatenated text content.
        string extractedText = plainText.Text;

        // Display the extracted text.
        Console.WriteLine(extractedText);
    }
}
