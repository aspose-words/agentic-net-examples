using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Input PDF file path
        string pdfPath = "InputDocument.pdf";

        // Output DOC file path
        string docPath = "ConvertedDocument.doc";

        // Load the PDF document using the Document constructor that accepts a file name
        Document pdfDoc = new Document(pdfPath);

        // Save the loaded document as a DOC file using the Save method with SaveFormat.Doc
        pdfDoc.Save(docPath, SaveFormat.Doc);
    }
}
