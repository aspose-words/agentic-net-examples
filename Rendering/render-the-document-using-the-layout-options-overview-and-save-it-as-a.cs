using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithLayoutOptions
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Example layout options: show hidden text and paragraph marks.
        doc.LayoutOptions.ShowHiddenText = true;
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Rebuild the page layout so that the changes take effect.
        doc.UpdatePageLayout();

        // Configure PDF save options – set the page layout for the PDF viewer.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PageLayout = PdfPageLayout.TwoPageLeft   // display two pages at a time, odd pages on the left
        };

        // Save the rendered document as PDF using the specified options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
