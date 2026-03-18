using System;
using Aspose.Words;

class PdfToEpubConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Path for the resulting EPUB file.
        string epubPath = "output.epub";

        // Load the PDF document using the Document constructor.
        Document doc = new Document(pdfPath);

        // Save the loaded document as an EPUB e‑book.
        doc.Save(epubPath, SaveFormat.Epub);
    }
}
