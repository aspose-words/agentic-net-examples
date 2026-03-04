using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\Sample.pdf";

        // Load the DOCX document (lifecycle rule: Document(string)).
        Document doc = new Document(inputPath);

        // Create and configure PDF save options (lifecycle rule: new PdfSaveOptions()).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render the PDF with high‑quality (slower) algorithms.
            UseHighQualityRendering = true,

            // Embed all fonts in the PDF so that the appearance is preserved on any device.
            EmbedFullFonts = true,

            // Configure the document outline (table of contents) that appears in PDF viewers.
            OutlineOptions = { HeadingsOutlineLevels = 3, ExpandedOutlineLevels = 2 },

            // Apply ZIP (Flate) compression to the textual content of the PDF.
            TextCompression = PdfTextCompression.Flate,

            // Ensure the PDF complies with the PDF/A‑1b archival standard.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as PDF using the custom options (lifecycle rule: Save(string, SaveOptions)).
        doc.Save(outputPath, pdfOptions);
    }
}
