using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("MyDir/Document.docx");

        // Create PDF save options to control the conversion.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A-1b compliance (preserves visual appearance).
            Compliance = PdfCompliance.PdfA1b,

            // Use high‑quality rendering algorithms (slower but better output).
            UseHighQualityRendering = true,

            // Export the document structure (tags) to the PDF.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF file using the specified options.
        doc.Save("ArtifactsDir/Document.pdf", pdfOptions);
    }
}
