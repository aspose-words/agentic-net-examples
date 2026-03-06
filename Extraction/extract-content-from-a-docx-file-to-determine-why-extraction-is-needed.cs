using System;
using Aspose.Words;

class ExtractPlainText
{
    static void Main()
    {
        // Path to the DOCX file to be extracted.
        string docxPath = @"C:\Path\To\Your\Document.docx";

        // Load the document as a plain text document.
        // The constructor automatically detects the file format.
        PlainTextDocument plainTextDoc = new PlainTextDocument(docxPath);

        // Retrieve the concatenated textual content.
        string extractedText = plainTextDoc.Text;

        // Output the extracted text to the console (or process as needed).
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);
    }
}
