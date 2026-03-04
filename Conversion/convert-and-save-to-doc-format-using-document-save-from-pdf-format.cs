using System;
using Aspose.Words;

class PdfToDocConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\source.pdf";

        // Path where the converted DOC file will be saved.
        string docPath = @"C:\Output\converted.doc";

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Save the loaded document as a legacy Word .doc file.
        pdfDocument.Save(docPath, SaveFormat.Doc);
    }
}
