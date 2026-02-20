using System;
using Aspose.Words;

class ExtractPlainText
{
    static void Main()
    {
        // Path to the DOCX file to extract text from.
        string docPath = @"C:\Docs\input.docx";

        // Load the document as a plain‑text representation.
        // The constructor automatically detects the file format.
        PlainTextDocument plainTextDoc = new PlainTextDocument(docPath);

        // Retrieve the concatenated textual content.
        string extractedText = plainTextDoc.Text;

        // Output the extracted text.
        Console.WriteLine(extractedText);
    }
}
