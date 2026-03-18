using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFilePath = "input.pdf";

        // Path where the extracted plain‑text will be saved.
        string txtFilePath = "output.txt";

        // Load the PDF document using the Document(string) constructor (load rule).
        Document pdfDocument = new Document(pdfFilePath);

        // Extract all textual content from the loaded document.
        string extractedText = pdfDocument.GetText();

        // Write the extracted text to a plain TXT file.
        File.WriteAllText(txtFilePath, extractedText);
    }
}
