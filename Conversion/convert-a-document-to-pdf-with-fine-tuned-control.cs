using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options to fine‑tune the conversion.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure the output complies with PDF/A‑1b (archival) standard.
            Compliance = PdfCompliance.PdfA1b,

            // Render the PDF with high‑quality (slower) algorithms.
            UseHighQualityRendering = true,

            // Embed all fonts so the PDF looks the same on any machine.
            EmbedFullFonts = true
        };

        // Configure the outline (bookmarks) that appears in PDF viewers.
        pdfOptions.OutlineOptions.HeadingsOutlineLevels = 3;   // include headings up to level 3
        pdfOptions.OutlineOptions.ExpandedOutlineLevels = 1; // expand only top‑level entries

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
