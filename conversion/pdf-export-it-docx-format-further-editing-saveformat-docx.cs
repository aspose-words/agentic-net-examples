using System;
using Aspose.Words;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfFile = "input.pdf";

        // Path for the resulting DOCX file.
        string docxFile = "output.docx";

        // Load the PDF document.
        Document pdfDocument = new Document(pdfFile);

        // Save the loaded document as DOCX.
        pdfDocument.Save(docxFile, SaveFormat.Docx);
    }
}
