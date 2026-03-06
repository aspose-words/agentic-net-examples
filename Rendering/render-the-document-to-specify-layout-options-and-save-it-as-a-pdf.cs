using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("MyDir/Big document.docx");

        // Specify layout options for rendering.
        // Example: show hidden text and paragraph marks in the output PDF.
        doc.LayoutOptions.ShowHiddenText = true;
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Rebuild the page layout after changing layout options.
        doc.UpdatePageLayout();

        // Create PDF save options and set the desired page layout.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Display two pages at a time, with odd-numbered pages on the left.
            PageLayout = PdfPageLayout.TwoPageLeft
        };

        // Save the document as a PDF using the configured options.
        doc.Save("ArtifactsDir/RenderedDocument.pdf", pdfOptions);
    }
}
