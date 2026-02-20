using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be processed.
        string docxPath = "input.docx";

        // Load the DOCX file as a plain‑text document.
        // The constructor automatically detects the format.
        PlainTextDocument plainText = new PlainTextDocument(docxPath);

        // Extract the concatenated textual content.
        string extractedText = plainText.Text;

        // Display the extracted text.
        Console.WriteLine(extractedText);
    }
}
