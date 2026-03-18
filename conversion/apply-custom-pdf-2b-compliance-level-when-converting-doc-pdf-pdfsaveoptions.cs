using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA2bConversion
{
    static void Main()
    {
        // Input DOC/DOCX file.
        string inputPath = @"MyDir\Document.docx";

        // Output PDF file that will comply with PDF/A‑2b.
        // In Aspose.Words the PDF/A‑2b level is represented by PdfCompliance.PdfA2u.
        string outputPath = @"ArtifactsDir\Document.PdfA2b.pdf";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Create save options for PDF output.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Apply PDF/A‑2b compliance.
        saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as PDF with the specified compliance level.
        doc.Save(outputPath, saveOptions);
    }
}
