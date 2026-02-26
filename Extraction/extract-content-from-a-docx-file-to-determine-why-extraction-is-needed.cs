using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be extracted.
        string docxPath = @"C:\Docs\Sample.docx";

        // Load the document as plain text using the PlainTextDocument constructor.
        PlainTextDocument plainText = new PlainTextDocument(docxPath);

        // Retrieve the concatenated text content.
        string extractedText = plainText.Text;

        // Output the extracted text to the console (or process as needed).
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);
    }
}
