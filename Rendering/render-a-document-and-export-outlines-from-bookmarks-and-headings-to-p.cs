using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportOutlinesToPdf
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = Path.Combine("MyDir", "Input.docx");

        // Path where the resulting PDF will be saved.
        string outputPath = Path.Combine("ArtifactsDir", "Output.pdf");

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Configure PDF save options to include outlines for bookmarks and headings.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Show the outline (bookmarks) pane when the PDF is opened.
        pdfOptions.PageMode = PdfPageMode.UseOutlines;
        // Display all Word bookmarks at the first level of the outline.
        pdfOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
        // Include headings up to level 3 in the outline (Heading1‑Heading3).
        pdfOptions.OutlineOptions.HeadingsOutlineLevels = 3;
        // Expand the first two outline levels automatically.
        pdfOptions.OutlineOptions.ExpandedOutlineLevels = 2;
        // Export bookmarks that are located in headers/footers (optional).
        pdfOptions.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.All;

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}
