using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be processed.
        string sourcePath = "input.docx";

        // Load the document using the Aspose.Words Document constructor (load rule).
        Document doc = new Document(sourcePath);

        // Extract the plain‑text representation of the document.
        // PlainTextDocument automatically detects the format and provides the concatenated text.
        PlainTextDocument plain = new PlainTextDocument(sourcePath);
        string extractedText = plain.Text;

        // Output the extracted text to the console.
        Console.WriteLine(extractedText);

        // Optionally, write the extracted text to a separate .txt file.
        System.IO.File.WriteAllText("extracted.txt", extractedText);
    }
}
