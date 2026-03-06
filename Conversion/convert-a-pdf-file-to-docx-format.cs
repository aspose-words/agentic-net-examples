using System;
using Aspose.Words;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Desired path for the converted DOCX file.
        string docxPath = "output.docx";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDocument = new Document(pdfPath);

        // Save the document as DOCX. The format is inferred from the .docx extension.
        pdfDocument.Save(docxPath);
    }
}
