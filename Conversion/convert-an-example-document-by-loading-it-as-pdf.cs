using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., DOCX).
        string sourcePath = "ExampleDocument.docx";

        // Path where the PDF will be saved.
        string pdfPath = "ExampleDocument.pdf";

        // Load the source document.
        Document doc = new Document(sourcePath);

        // Save the loaded document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
