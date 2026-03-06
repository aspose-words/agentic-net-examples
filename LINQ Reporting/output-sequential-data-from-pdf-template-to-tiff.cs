using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToTiffConverter
{
    static void Main()
    {
        // Load the PDF template.
        Document doc = new Document("Template.pdf");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Ensure all pages are saved. In older Aspose.Words versions the MultiPageLayout
        // property does not exist; the document is saved as a multi‑frame TIFF when the
        // PageSet includes the whole document.
        saveOptions.PageSet = new PageSet(0, doc.PageCount - 1);

        // Save the document as a multi‑frame TIFF image.
        doc.Save("Output.tiff", saveOptions);
    }
}
