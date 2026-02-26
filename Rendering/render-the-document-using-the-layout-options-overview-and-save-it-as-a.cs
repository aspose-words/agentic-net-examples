using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithLayoutOptions
{
    static void Main()
    {
        // Paths to the source document and the output PDF.
        string dataDir = @"C:\Data\";
        string artifactsDir = @"C:\Artifacts\";

        // Load the source Word document.
        Document doc = new Document(dataDir + "Input.docx");

        // Enable layout options that affect rendering.
        // Show hidden text and paragraph marks in the rendered PDF.
        doc.LayoutOptions.ShowHiddenText = true;
        doc.LayoutOptions.ShowParagraphMarks = true;

        // Rebuild the page layout so that the changes to LayoutOptions take effect.
        doc.UpdatePageLayout();

        // Create PDF save options and set a page layout for the PDF viewer.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PageLayout = PdfPageLayout.TwoPageLeft   // display two pages at a time, odd pages on the left
        };

        // Save the document as a PDF using the specified options.
        doc.Save(artifactsDir + "RenderedDocument.pdf", pdfOptions);
    }
}
