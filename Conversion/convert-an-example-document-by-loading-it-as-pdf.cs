using System;
using System.IO;
using Aspose.Words;

class PdfConversionExample
{
    static void Main()
    {
        // Path to the source document (can be any supported format, e.g., .docx, .pdf, etc.).
        string inputFile = @"C:\Input\ExampleDocument.pdf";

        // Path where the converted PDF will be saved.
        // Saving with a .pdf extension will keep the format as PDF.
        string outputFile = @"C:\Output\ConvertedDocument.pdf";

        // Load the document. The constructor automatically detects the format (PDF in this case).
        Document doc = new Document(inputFile);

        // Save the loaded document as PDF.
        doc.Save(outputFile);
    }
}
