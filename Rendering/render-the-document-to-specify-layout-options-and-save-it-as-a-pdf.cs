using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class RenderAndSavePdf
{
    static void Main()
    {
        // Paths to the input Word document and the output PDF file.
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Output\RenderedDocument.pdf";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Rebuild the page layout to ensure that any layout‑dependent options are up‑to‑date.
        doc.UpdatePageLayout();

        // Specify layout options for the rendered output.
        // Show hidden text and paragraph marks in the PDF.
        doc.LayoutOptions.ShowHiddenText = true;
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Create PDF save options and configure the page layout that the PDF viewer will use.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Display two pages side‑by‑side, with odd‑numbered pages on the left.
            PageLayout = PdfPageLayout.TwoPageLeft,

            // Show the document outline (bookmarks) when the PDF is opened.
            PageMode = PdfPageMode.UseOutlines,

            // Export the document structure (tags) to aid accessibility tools.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF using the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
