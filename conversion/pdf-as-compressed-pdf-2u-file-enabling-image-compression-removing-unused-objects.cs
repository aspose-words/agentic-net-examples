using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document (replace with your actual file path)
        var doc = new Document("Input.docx");

        // Configure PDF save options for PDF/A‑2u with image compression and output optimisation
        var pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A‑2u compliance
            Compliance = PdfCompliance.PdfA2u,

            // Enable JPEG image compression (quality 0‑100)
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 50,

            // Remove unused objects and redundant canvases to reduce file size
            OptimizeOutput = true
        };

        // Save the document as a compressed PDF/A‑2u file
        doc.Save("Output_Compressed_PdfA2u.pdf", pdfOptions);
    }
}
