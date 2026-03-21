using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple Word document programmatically
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a sample PDF/A‑2u document with image compression.");

        // Configure PDF save options for PDF/A‑2u with image compression and output optimisation
        var pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A‑2u compliance
            Compliance = PdfCompliance.PdfA2u,

            // Enable JPEG image compression (quality 0‑100)
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 50,

            // Remove unused objects and redundant canvases to reduce file size
            OptimizeOutput = true,

            // Do not update fields during save to avoid processing unsupported fields
            UpdateFields = false
        };

        // Save the document as a compressed PDF/A‑2u file
        doc.Save("Output_Compressed_PdfA2u.pdf", pdfOptions);
    }
}
