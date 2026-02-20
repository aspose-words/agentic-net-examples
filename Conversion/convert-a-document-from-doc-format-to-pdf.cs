using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file. The constructor automatically detects the format.
        Document document = new Document("input.doc");

        // Create PDF save options (optional – can be customized here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF using the specified options.
        document.Save("output.pdf", pdfOptions);
    }
}
