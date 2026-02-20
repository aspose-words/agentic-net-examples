using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Example.docx");

        // Create PDF save options – default settings are sufficient for a basic conversion.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save("Example.pdf", pdfOptions);
    }
}
