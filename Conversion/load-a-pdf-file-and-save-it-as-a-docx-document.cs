using System;
using Aspose.Words;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFilePath = @"C:\Path\To\Input.pdf";

        // Path where the converted DOCX file will be saved.
        string docxFilePath = @"C:\Path\To\Output.docx";

        // Load the PDF document using the Document constructor that accepts a file name.
        Document pdfDocument = new Document(pdfFilePath);

        // Save the loaded document as DOCX.
        // The Save method automatically determines the format from the file extension.
        pdfDocument.Save(docxFilePath);
    }
}
