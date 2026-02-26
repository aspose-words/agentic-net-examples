using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderToPdf
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PDF save options and configure them.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export the document structure (tags) to the PDF.
            ExportDocumentStructure = true
        };

        // Save the document as PDF using the specified options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
