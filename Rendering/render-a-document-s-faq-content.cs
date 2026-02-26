using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderFaq
{
    static void Main()
    {
        // Path to the source FAQ document (DOCX, DOC, etc.).
        string inputFile = @"C:\Docs\FaqDocument.docx";

        // Path where the rendered PDF will be saved.
        string outputFile = @"C:\Docs\FaqDocument.pdf";

        // Load the existing document.
        Document doc = new Document(inputFile);

        // Create PDF save options and configure DrawingML rendering.
        // Use DrawingML to render the original shapes; change to Fallback if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML
        };

        // Save the document as PDF using the specified options.
        doc.Save(outputFile, pdfOptions);
    }
}
