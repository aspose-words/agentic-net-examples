using System;
using Aspose.Words;
using Aspose.Words.Loading;

namespace PdfToDocxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file that contains form fields.
            string pdfPath = @"C:\Input\form_fields.pdf";

            // Path where the resulting DOCX file will be saved.
            string docxPath = @"C:\Output\form_fields_converted.docx";

            // Load the PDF document. Aspose.Words automatically detects the format.
            // No special load options are required for preserving form fields;
            // they are imported as Word form fields (StructuredDocumentTag objects).
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded document as DOCX. The file extension determines the save format.
            pdfDocument.Save(docxPath);
        }
    }
}
