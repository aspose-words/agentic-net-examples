using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToPdfWithFineControl
{
    static void Main()
    {
        // Load the source document from a file using the provided Document(string) constructor.
        Document doc = new Document("InputDocuments/SourceDocument.docx");

        // Create a PdfSaveOptions object – this is the dedicated options class for PDF output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Fine‑tuned settings:

        // 1. Ensure the PDF complies with PDF/A‑2u (preserves visual appearance and text extraction).
        pdfOptions.Compliance = PdfCompliance.PdfA2u;

        // 2. Reduce memory usage for very large documents.
        pdfOptions.MemoryOptimization = true;

        // 3. Use high‑quality rendering (slower but better visual fidelity).
        pdfOptions.UseHighQualityRendering = true;

        // 4. Enable anti‑aliasing for smoother text and graphics.
        pdfOptions.UseAntiAliasing = true;

        // 5. Configure the outline (bookmarks) – include only headings up to level 3.
        pdfOptions.OutlineOptions.HeadingsOutlineLevels = 3;
        // Expand the first two heading levels when the PDF is opened.
        pdfOptions.OutlineOptions.ExpandedOutlineLevels = 2;

        // 6. Set the page layout to display two pages side‑by‑side.
        pdfOptions.PageLayout = PdfPageLayout.TwoColumnLeft;

        // 7. Optionally, embed all fonts to avoid missing glyphs on other machines.
        pdfOptions.EmbedFullFonts = true;

        // Save the document as PDF using the overload that accepts a file name and SaveOptions.
        doc.Save("OutputDocuments/ResultDocument.pdf", pdfOptions);
    }
}
