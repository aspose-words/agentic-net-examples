using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderFirstPageToPng
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Configure image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Render only the first page (zero‑based index).
        options.PageSet = new PageSet(0);

        // Save the rendered page as a PNG image.
        doc.Save("FirstPage.png", options);
    }
}
