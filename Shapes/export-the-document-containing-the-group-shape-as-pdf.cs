using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportGroupShapeToPdf
{
    static void Main()
    {
        // Load the Word document that contains a GroupShape.
        Document doc = new Document("GroupShapeDocument.docx");

        // Create PDF save options (default settings are sufficient for rendering the group shape).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF; the GroupShape will be rendered automatically.
        doc.Save("GroupShapeDocument.pdf", pdfOptions);
    }
}
