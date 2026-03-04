using System;
using Aspose.Words;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path where the converted DOCX file will be saved.
        string docxPath = @"C:\Output\sample.docx";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDocument = new Document(pdfPath);

        // Save the document as DOCX. The file extension determines the save format.
        pdfDocument.Save(docxPath);
    }
}
